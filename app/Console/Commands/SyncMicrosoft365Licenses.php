<?php

namespace App\Console\Commands;

use App\Models\License;
use App\Models\User;
use Illuminate\Console\Command;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;

class SyncMicrosoft365Licenses extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'snipeit:sync-m365-licenses
                            {--license_id= : The Snipe-IT license ID to assign}
                            {--sku= : The Microsoft 365 SKU ID to filter (e.g., SPE_E5 or the GUID)}
                            {--tenant_id= : Azure AD Tenant ID}
                            {--client_id= : Azure AD App Client ID}
                            {--client_secret= : Azure AD App Client Secret}
                            {--match_by=email : How to match users (email, username, employee_num)}
                            {--notify : Send notification emails}
                            {--dry-run : Show what would be done without making changes}';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Sync Microsoft 365 E5 licenses to Snipe-IT users';

    /**
     * Common Microsoft 365 E5 SKU IDs
     */
    protected array $e5SkuIds = [
        '06ebc4ee-1bb5-47dd-8120-11324bc54e06', // Microsoft 365 E5
        'c7df2760-2c81-4ef7-b578-5b5392b571df', // Microsoft 365 E5 (without Audio Conferencing)
        'a91fc4e0-65e5-4266-aa76-4037509c1626', // Microsoft 365 E5 Compliance
        '184efa21-98c3-4e5d-95ab-d07053a96e67', // Microsoft 365 E5 Security
    ];

    /**
     * Execute the console command.
     */
    public function handle(): int
    {
        $licenseId = $this->option('license_id');
        $skuFilter = $this->option('sku');
        $tenantId = $this->option('tenant_id') ?? config('services.microsoft.tenant_id') ?? env('MS365_TENANT_ID');
        $clientId = $this->option('client_id') ?? config('services.microsoft.client_id') ?? env('MS365_CLIENT_ID');
        $clientSecret = $this->option('client_secret') ?? config('services.microsoft.client_secret') ?? env('MS365_CLIENT_SECRET');
        $matchBy = $this->option('match_by');
        $notify = $this->option('notify');
        $dryRun = $this->option('dry-run');

        // Validate required parameters
        if (!$licenseId) {
            $this->error('ERROR: --license_id is required.');
            return Command::FAILURE;
        }

        if (!$tenantId || !$clientId || !$clientSecret) {
            $this->error('ERROR: Azure AD credentials are required. Provide via options or environment variables (MS365_TENANT_ID, MS365_CLIENT_ID, MS365_CLIENT_SECRET).');
            return Command::FAILURE;
        }

        // Get the Snipe-IT license
        $license = License::with('assignedusers')->find($licenseId);
        if (!$license) {
            $this->error('ERROR: License with ID ' . $licenseId . ' not found.');
            return Command::FAILURE;
        }

        $this->info('Syncing Microsoft 365 licenses to Snipe-IT license: ' . $license->name);

        if ($dryRun) {
            $this->warn('DRY RUN MODE - No changes will be made.');
        }

        // Get access token from Microsoft
        $accessToken = $this->getAccessToken($tenantId, $clientId, $clientSecret);
        if (!$accessToken) {
            $this->error('ERROR: Failed to authenticate with Microsoft Graph API.');
            return Command::FAILURE;
        }

        $this->info('Successfully authenticated with Microsoft Graph API.');

        // Get users with E5 licenses from Microsoft 365
        $m365Users = $this->getMicrosoft365UsersWithLicense($accessToken, $skuFilter);
        if ($m365Users === null) {
            $this->error('ERROR: Failed to retrieve users from Microsoft Graph API.');
            return Command::FAILURE;
        }

        $this->info('Found ' . count($m365Users) . ' users with E5 license in Microsoft 365.');

        $assigned = 0;
        $skipped = 0;
        $notFound = 0;
        $noSeats = 0;

        foreach ($m365Users as $m365User) {
            $snipeUser = $this->findSnipeUser($m365User, $matchBy);

            if (!$snipeUser) {
                $this->warn('User not found in Snipe-IT: ' . ($m365User['mail'] ?? $m365User['userPrincipalName']));
                $notFound++;
                continue;
            }

            // Check if user already has this license
            if ($snipeUser->licenses->where('id', '=', $licenseId)->count() > 0) {
                $this->line('User ' . $snipeUser->username . ' already has this license. Skipping...');
                $skipped++;
                continue;
            }

            // Check available seats
            if ($license->availCount()->count() < 1) {
                $this->error('ERROR: No available license seats remaining.');
                $noSeats++;
                continue;
            }

            if ($dryRun) {
                $this->info('[DRY RUN] Would assign license to: ' . $snipeUser->username . ' (' . $snipeUser->email . ')');
                $assigned++;
                continue;
            }

            // Get a free seat and assign it
            $licenseSeat = $license->freeSeat();
            if (!$licenseSeat) {
                $this->error('ERROR: Could not get a free license seat.');
                $noSeats++;
                continue;
            }

            $licenseSeat->assigned_to = $snipeUser->id;

            if ($licenseSeat->save()) {
                // Handle notification
                if (!$notify) {
                    $originalEmail = $snipeUser->email;
                    $snipeUser->email = null;
                }

                $licenseSeat->logCheckout('Synced from Microsoft 365 E5 license', $snipeUser);

                if (!$notify) {
                    $snipeUser->email = $originalEmail;
                }

                $this->info('License assigned to: ' . $snipeUser->username);
                $assigned++;

                // Refresh the license to get updated seat count
                $license->refresh();
            } else {
                $this->error('ERROR: Failed to save license seat for user: ' . $snipeUser->username);
            }
        }

        $this->newLine();
        $this->info('=== Sync Summary ===');
        $this->info('Assigned: ' . $assigned);
        $this->info('Already had license: ' . $skipped);
        $this->info('Not found in Snipe-IT: ' . $notFound);
        if ($noSeats > 0) {
            $this->warn('Skipped due to no seats: ' . $noSeats);
        }

        return Command::SUCCESS;
    }

    /**
     * Get an access token from Microsoft Graph API
     */
    protected function getAccessToken(string $tenantId, string $clientId, string $clientSecret): ?string
    {
        try {
            $response = Http::asForm()->post(
                "https://login.microsoftonline.com/{$tenantId}/oauth2/v2.0/token",
                [
                    'client_id' => $clientId,
                    'client_secret' => $clientSecret,
                    'scope' => 'https://graph.microsoft.com/.default',
                    'grant_type' => 'client_credentials',
                ]
            );

            if ($response->successful()) {
                return $response->json('access_token');
            }

            $this->error('Token request failed: ' . $response->body());
            Log::error('Microsoft Graph token request failed', ['response' => $response->body()]);
            return null;
        } catch (\Exception $e) {
            $this->error('Token request exception: ' . $e->getMessage());
            Log::error('Microsoft Graph token request exception', ['error' => $e->getMessage()]);
            return null;
        }
    }

    /**
     * Get users with E5 licenses from Microsoft 365
     */
    protected function getMicrosoft365UsersWithLicense(string $accessToken, ?string $skuFilter = null): ?array
    {
        $skuIdsToMatch = $this->e5SkuIds;

        // If a specific SKU filter is provided
        if ($skuFilter) {
            if ($this->isGuid($skuFilter)) {
                $skuIdsToMatch = [$skuFilter];
            } else {
                // Try to look up the SKU by name (optional - may fail without Organization.Read.All)
                $skuResponse = Http::withToken($accessToken)
                    ->get('https://graph.microsoft.com/v1.0/subscribedSkus');

                if ($skuResponse->successful()) {
                    foreach ($skuResponse->json('value', []) as $sku) {
                        if (stripos($sku['skuPartNumber'], $skuFilter) !== false) {
                            $skuIdsToMatch = [$sku['skuId']];
                            $this->info('Found SKU: ' . $sku['skuPartNumber'] . ' (' . $sku['skuId'] . ')');
                            break;
                        }
                    }
                } else {
                    $this->warn('Could not look up SKU by name (needs Organization.Read.All). Using default E5 SKU IDs.');
                }
            }
        }

        $this->info('Looking for users with SKU IDs: ' . implode(', ', $skuIdsToMatch));

        // Directly fetch users with assignedLicenses
        return $this->getMicrosoft365UsersWithLicenseAlternative($accessToken, $skuIdsToMatch);
    }

    /**
     * Alternative method to get users - fetches all users and checks licenses individually
     */
    protected function getMicrosoft365UsersWithLicenseAlternative(string $accessToken, array $skuIdsToMatch): ?array
    {
        $users = [];
        $this->info('Using alternative method: fetching all users and checking licenses individually...');

        try {
            $nextLink = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,userPrincipalName,employeeId,assignedLicenses&$top=999';

            while ($nextLink) {
                $response = Http::withToken($accessToken)->get($nextLink);

                if (!$response->successful()) {
                    $this->error('Failed to get users: ' . $response->body());
                    Log::error('Microsoft Graph users request failed', ['response' => $response->body()]);
                    return null;
                }

                $data = $response->json();

                foreach ($data['value'] ?? [] as $user) {
                    $assignedLicenses = $user['assignedLicenses'] ?? [];

                    foreach ($assignedLicenses as $license) {
                        if (in_array($license['skuId'], $skuIdsToMatch)) {
                            $users[] = [
                                'id' => $user['id'],
                                'displayName' => $user['displayName'],
                                'mail' => $user['mail'],
                                'userPrincipalName' => $user['userPrincipalName'],
                                'employeeId' => $user['employeeId'] ?? null,
                                'skuId' => $license['skuId'],
                            ];
                            break;
                        }
                    }
                }

                $nextLink = $data['@odata.nextLink'] ?? null;
            }

            return $users;
        } catch (\Exception $e) {
            $this->error('Exception while fetching users (alternative): ' . $e->getMessage());
            Log::error('Microsoft Graph users request exception', ['error' => $e->getMessage()]);
            return null;
        }
    }

    /**
     * Find a Snipe-IT user based on Microsoft 365 user data
     */
    protected function findSnipeUser(array $m365User, string $matchBy): ?User
    {
        $query = User::whereNull('deleted_at');

        switch ($matchBy) {
            case 'email':
                $email = $m365User['mail'] ?? $m365User['userPrincipalName'];
                return $query->where('email', '=', $email)
                    ->orWhere('email', '=', $m365User['userPrincipalName'])
                    ->with('licenses')
                    ->first();

            case 'username':
                // Extract username from UPN (before @)
                $upn = $m365User['userPrincipalName'];
                $username = strstr($upn, '@', true) ?: $upn;
                return $query->where('username', '=', $username)
                    ->with('licenses')
                    ->first();

            case 'employee_num':
                if (empty($m365User['employeeId'])) {
                    return null;
                }
                return $query->where('employee_num', '=', $m365User['employeeId'])
                    ->with('licenses')
                    ->first();

            default:
                // Try email first, then username
                $email = $m365User['mail'] ?? $m365User['userPrincipalName'];
                $upn = $m365User['userPrincipalName'];
                $username = strstr($upn, '@', true) ?: $upn;

                return $query->where('email', '=', $email)
                    ->orWhere('email', '=', $upn)
                    ->orWhere('username', '=', $username)
                    ->with('licenses')
                    ->first();
        }
    }

    /**
     * Check if a string is a valid GUID
     */
    protected function isGuid(string $string): bool
    {
        return preg_match('/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i', $string) === 1;
    }
}
