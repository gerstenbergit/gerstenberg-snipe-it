@props([
    'count' => null,
])

@can('view', \App\Models\Accessory::class)
    <x-tabs.nav-item
            name="accessories"
            icon_type="accessory"
            label="{{ trans('general.accessories') }}"
            count="{{ $count }}"
            tooltip="{{ trans('general.accessories') }}"
    />
@endcan