@props([
    'count' => null,
])

@can('view', \App\Models\Asset::class)
    <x-tabs.nav-item
            name="assets"
            class="active"
            icon_type="assets"
            label="{{ trans('general.assets') }}"
            count="{{ $count }}"
            tooltip="{{ trans('general.assets') }}"
    />
@endcan