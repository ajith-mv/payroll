@extends('layouts.template')
@section('content')
    <style>
        .service-pagination nav {
            display: flex;
            justify-content: space-between;
            align-content: center;
            align-items: center;
        }
    </style>
    <section style="min-height: 93vh" class="p-3">
        <div class="card shadow border border-secondary rounded">
            <div class="bg-light border-bottom p-2 d-flex align-items-center justify-content-between">
                <b>Service Book History Report</b>

                <form action="{{ route('reports.service.history') }}" class="input-group w-auto d-inline" method="GET">
                    <div class="d-flex">
                        <div>
                            <a href="{{ route('reports.service.history') }}" class="btn btn-sm btn-warning"><i
                                    class="fa fa-repeat"></i> Reset</a>
                        </div>
                        <div>

                            <select name="employee" id="employee" class="form-control">
                                <option value="">-- Select Employee --</option>
                                @foreach ($employees as $items)
                                    <option value="{{ $items->id }}"
                                        @if ($employee_id == $items->id) selected="selected" @endif>
                                        {{ $items->name }} - {{ $items->society_emp_code }}</option>
                                @endforeach
                            </select>
                        </div>
                        <div>

                            <select name="department" class="form-select form-select-sm w-auto d-inline">
                                <option value="">-- Select Department --</option>
                                @if( isset( $departments ) && !empty( $departments ) ) 
                                    @foreach ($departments as $item)
                                        <option value="{{ $item->id }}"  @if ($department_id == $item->id) selected="selected" @endif>{{ $item->name }}</option>
                                    @endforeach
                                @endif
                            </select>
                        </div>
                        <div>

                            <button onclick="this.form.action = '{{ route('reports.service.history') }}';" type="submit"
                                class="btn btn-sm btn-primary"><i class="fa fa-search"></i></button>
                            {{-- <a href="{{ route('reports.service.history.export') }}" class="btn btn-sm btn-success"><i class="fa fa-print me-2"></i>Pdf</button> --}}
                            {{-- <button type="button" onclick="printServiceAreaHistory('print_service_area')" class="btn btn-sm btn-success"><i class="fa fa-print me-2"></i>Print</button> --}}
                            <input type="submit"
                                onclick="this.form.action = '{{ route('reports.service.history.export') }}';"
                                class="btn btn-sm btn-danger" formtarget="_blank" value="Pdf" />
                        </div>


                    </div>
                </form>
            </div>
            <div class="card-body p-2">
                <div class="table-responsive" id="print_service_area">
                    @include('pages.reports.service_history._list_card')
                </div>

            </div>
            <div class="card-footer p-2 bg-light service-pagination">
                {!! $paginate_link !!}
            </div>
        </div>
    </section>
    <script>
        $('#employee').select2({
            theme: 'bootstrap-5'
        });
    </script>
@endsection
