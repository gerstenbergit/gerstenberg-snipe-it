@props([
    'count' => null,
    'model' => null,
])


@can('view', $model)
    <x-tabs.nav-item
            name="history"
            icon_type="history"
            label="{{ trans('general.history') }}"
            count="{{ $count }}"
            tooltip="{{ trans('general.history') }}"
    />
@endcan