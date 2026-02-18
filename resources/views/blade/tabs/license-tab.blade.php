@props([
    'count' => null,
])

@can('view', \App\Models\License::class)
    <x-tabs.nav-item
            name="licenses"
            icon_type="licenses"
            label="{{ trans('general.licenses') }}"
            count="{{ $count }}"
            tooltip="{{ trans('general.licenses') }}"
    />
@endcan