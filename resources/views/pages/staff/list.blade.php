@extends('layouts.template')
@section('breadcrum')
    @include('layouts.parts.breadcrum')
@endsection
@section('content')
    <style>
        #staff_table td {
            padding-left: 10px;
        }

        #staff_table_filter {
            display: none;
        }
      
    </style>
    <div class="card">

        <div class="card-header border-0 pt-6">

            <div class="card-title">

                <div class="d-flex align-items-center position-relative my-1">
                    <div>
                        <span class="mt-2">

                            {!! searchSvg() !!}
                        </span>
                        <input type="text" name="datatable_search" data-kt-user-table-filter="search"
                            id="staff_datable_search" class="form-control  w-250px ps-14"
                            placeholder="Search user">
                    </div>&nbsp;&nbsp; 

                    <div class="form-group">
                        <select name="verification_status" id="verification_status" class="form-control">
                            <option value=""> All Profile Status </option>
                            <option value="pending">Pending</option>
                            <option value="approved">Approved</option>
                        </select>
                    </div>
&nbsp;&nbsp;      <div class="form-group">
            <input type="number" name="page_search"
                            id="page_search" class="form-control  w-250px ps-14"
                            placeholder="Go to Page" value="1">
                    </div>

                    {{-- <div class="form-group">
                        <select name="datatable_institute_id" id="datatable_institute_id" class="form-control">
                            <option value=""> All Institution </option>
                            @if (isset($institutions) && !empty($institutions))
                                @foreach ($institutions as $item)
                                    <option value="{{ $item->id }}"> {{ $item->name }}</option>
                                @endforeach
                            @endif
                        </select>
                    </div> --}}
                </div>

            </div>
            @php
                $route_name = request()
                    ->route()
                    ->getName();
                
            @endphp

            @if (access()->buttonAccess($route_name, 'add_edit'))
                <div class="card-toolbar">
                    <div class="d-flex justify-content-end">
                        <a href="{{ route('staff.register') }}" class="btn btn-primary btn-sm">
                            {!! plusSvg() !!}
                            Add User
                        </a>
                    </div>
                </div>
            @endif

        </div>

        <div class="card-body py-4">
            <div id="kt_table_users_wrapper" class="dataTables_wrapper dt-bootstrap4 no-footer">
                <div class="table-responsive">
                    <table class="table align-middle table-hover table-bordered table-striped fs-7 no-footer"
                        id="staff_table">
                        <thead class="bg-primary">
                            <tr class="text-start text-center text-muted fw-bolder fs-7 text-uppercase gs-0">

                                <th class="text-white text-start ps-3">
                                    Staff Info
                                </th>
                                <th class="text-white text-start ps-3">
                                    Society Code
                                </th>
                                <th class="text-white text-start ps-3">
                                    Institution Code
                                </th>
                                <th class="text-white text-start ps-3">
                                    Profile Completion
                                </th>
                                <th class="text-white">
                                    Status
                                </th>
                                <th class="text-white">
                                    Actions
                                </th>
                            </tr>
                        </thead>

                        <tbody class="text-gray-600 fw-bold">
                        </tbody>
                    </table>
                </div>

            </div>
        </div>
    </div>
@endsection

@section('add_on_script')
    <script> 
$('#page_search').val(1);
        var staff_Table = $('#staff_table').DataTable({
            processing: true,
            serverSide: true,
            order: [[1, "DESC"]],
            
            type: 'POST',
            "ajax": {
                "url": "{{ route('staff.list') }}",
                "dataType": "json",
                "data": function(d) {
                    d._token = "{{ csrf_token() }}",
                        d.staff_datable_search = $('#staff_datable_search').val(),
                        d.verification_status = $('#verification_status').val(),
                        // d.start = (($('#page_search').val() * 25),
                        d.datatable_institute_id = $('#datatable_institute_id').val()

                }
            },
            "columns": [{
                    "data": "name"
                },
                {
                    "data": "society_code"
                },
                {
                    "data": "institute_code"
                },
                {
                    "data": "profile"
                },
                {
                    "data": "status"
                },
                {
                    "data": "actions",
                    orderable: false,
                    searchable: false
                }
            ],
        language: {
            paginate: {
                next: '<i class="fa fa-angle-right"></i>', // or '→'
                previous: '<i class="fa fa-angle-left"></i>' // or '←' 
            }
        },
            "aLengthMenu": [
                [25, 50, 100, 200, 500, -1],
                [25, 50, 100, 200, 500, "All"]
            ]

        });
        
        document.querySelector('#staff_datable_search').addEventListener("keyup", function(e) {
            staff_Table.ajax.reload();
        });
        document.querySelector('#verification_status').addEventListener("change", function(e) {
            staff_Table.ajax.reload();
        });
         document.querySelector('#page_search').addEventListener("change", function(e) {
        var no=$("#page_search").val()-1;
       staff_Table.page(no).draw( 'page' );
            //staff_Table.ajax.reload();
        });
        document.querySelector('#datatable_institute_id').addEventListener("change", function(e) {
            staff_Table.ajax.reload();
        });
       staff_Table.ajax.reload();
        function staffChangeStatus(id, status) {
            Swal.fire({
                text: "Are you sure you would like to change status?",
                icon: "warning",
                showCancelButton: true,
                buttonsStyling: false,
                confirmButtonText: "Yes, Change it!",
                cancelButtonText: "No, return",
                customClass: {
                    confirmButton: "btn btn-danger",
                    cancelButton: "btn btn-active-light"
                }
            }).then(function(result) {
                if (result.value) {
                    $.ajaxSetup({
                        headers: {
                            'X-CSRF-TOKEN': $('meta[name="csrf-token"]').attr('content')
                        }
                    });

                    $.ajax({
                        url: "{{ route('staff.change.status') }}",
                        type: 'POST',
                        data: {
                            id: id,
                            status: status
                        },
                        success: function(res) {
                            staff_Table.ajax.reload();
                            Swal.fire({
                                title: "Updated!",
                                text: res.message,
                                icon: "success",
                                confirmButtonText: "Ok, got it!",
                                customClass: {
                                    confirmButton: "btn btn-success"
                                },
                                timer: 3000
                            });

                        },
                        error: function(xhr, err) {
                            if (xhr.status == 403) {
                                toastr.error(xhr.statusText, 'UnAuthorized Access');
                            }
                        }
                    });
                }
            });
        }

        function deleteStaff(id) {

            Swal.fire({
                text: "Are you sure you would like to delete?",
                icon: "warning",
                showCancelButton: true,
                buttonsStyling: false,
                confirmButtonText: "Yes, Delete it!",
                cancelButtonText: "No, return",
                customClass: {
                    confirmButton: "btn btn-danger",
                    cancelButton: "btn btn-active-light"
                }
            }).then(function(result) {
                if (result.value) {
                    $.ajaxSetup({
                        headers: {
                            'X-CSRF-TOKEN': $('meta[name="csrf-token"]').attr('content')
                        }
                    });

                    $.ajax({
                        url: "{{ route('staff.delete') }}",
                        type: 'POST',
                        data: {
                            id: id
                        },
                        success: function(res) {
                            staff_Table.ajax.reload();
                            Swal.fire({
                                title: "Deleted!",
                                text: res.message,
                                icon: "success",
                                confirmButtonText: "Ok, got it!",
                                customClass: {
                                    confirmButton: "btn btn-success"
                                },
                                timer: 3000
                            });

                        },
                        error: function(xhr, err) {
                            if (xhr.status == 403) {
                                toastr.error(xhr.statusText, 'UnAuthorized Access');
                            }
                        }
                    });
                }
            });
        }
    </script>
@endsection
