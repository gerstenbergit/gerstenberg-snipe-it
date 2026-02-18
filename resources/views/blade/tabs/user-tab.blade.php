@props([
    'count' => null,
])

@can('view', \App\Models\User::class)
    <x-tabs.nav-item
            name="users"
            icon_type="user"
            label="{{ trans('general.users') }}"
            count="{{ $count }}"
            tooltip="{{ trans('general.users') }}"
    />
@endcan