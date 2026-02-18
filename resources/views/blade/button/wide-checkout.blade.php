@props([
    'item' => null,
    'route' => null,
])

@can('checkout', $item)
    <a href="{{ $route  }}" class="btn btn-sm bg-maroon btn-social btn-block hidden-print">
        <x-icon type="checkout" />
        {{ trans('general.checkout') }}
    </a>
@endcan
