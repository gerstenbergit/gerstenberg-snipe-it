@props([
    'count' => null,
])
<x-tabs.nav-item
        name="files"
        icon_type="files"
        label="{{ trans('general.files') }}"
        count="{{ $count }}"
/>