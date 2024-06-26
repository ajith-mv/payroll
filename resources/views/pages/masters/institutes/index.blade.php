<!--begin::Navbar-->
@extends('layouts.template')
@section('breadcrum')
    @include('layouts.parts.breadcrum')
@endsection
@section('content')
<style>
    #institution_table td {
        padding-left: 10px;
        padding-right: 3px;
    }
</style>
    <div class="card">
        <div class="card-header border-0 pt-6">
            <div class="card-title">
                <div class="d-flex align-items-center position-relative my-1">
                    {!! searchSvg() !!}
                    <input type="text" data-kt-user-table-filter="search" id="institution_datable_search"
                        class="form-control form-control-solid w-250px ps-14" placeholder="Search Institutions">
                </div>
            </div>
            <div class="card-toolbar">
                <div class="d-flex justify-content-end" data-kt-user-table-toolbar="base">
                @php
                    $route_name = request()->route()->getName();               
                @endphp
                @if( access()->buttonAccess($route_name,'export') )
                    <a type="button" class="btn btn-light-primary me-3 btn-sm" href="{{ route('institutions.export') }}">
                        {!! exportSvg() !!} 
                        Export
                    </a>
                @endif
                @if( access()->buttonAccess($route_name,'add_edit') )
                    <button type="button" class="btn btn-primary btn-sm" id="add_modal" onclick="getInstituteModal()">
                        {!! plusSvg() !!} Add Institutions
                    </button>
                @endif

                </div>

                <div class="d-flex justify-content-end align-items-center d-none" data-kt-user-table-toolbar="selected">
                    <div class="fw-bolder me-5">
                        <span class="me-2" data-kt-user-table-select="selected_count"></span>Selected
                    </div>
                    <button type="button" class="btn btn-danger" data-kt-user-table-select="delete_selected">Delete
                        Selected
                    </button>
                </div>
            </div>
        </div>

        <div class="card-body py-4">
            <div id="kt_table_users_wrapper" class="dataTables_wrapper dt-bootstrap4 no-footer">
                <div class="table-responsive">
                    <table class="table align-middle  table-hover table-bordered table-striped fs-7 no-footer"
                        id="institution_table">
                        <thead class="bg-primary">
                            <tr class="text-start text-muted fw-bolder fs-7 text-uppercase gs-0">
                                
                                <th class="px-3 text-white" >
                                    Name
                                </th>
                                <th class="px-3 text-white" >
                                    Code
                                </th>
                                <th class="px-3 text-white" >
                                    Society
                                </th>
                                <th class="px-3 text-white" >
                                    Address
                                </th>
                                <th class="px-3 text-white">
                                    Last Updated
                                </th>

                                <th class="px-3 text-white">
                                    Status
                                </th>
                                <th class="px-3 text-white">
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
        var dtTable = $('#institution_table').DataTable({

            processing: true,
            serverSide: true,
            order :[0, 'desc'],
            type: 'POST',
            ajax: {
                "url": "{{ route('institutions') }}",
                "data": function(d) {
                    d.datatable_search = $('#institution_datable_search').val();
                }
            },

            columns: [
               
                {
                    data: 'name',
                    name: 'name'
                },

                {
                    data: 'code',
                    name: 'code'
                },
                {
                    data: 'society',
                    data: 'society'
                },
                {
                    data: 'address',
                    data: 'address'
                },
                {
                    data: 'last_updated',
                    data: 'last_updated',
                    searchable: false
                },
                {
                    data: 'status',
                    name: 'status'
                },
               
                {
                    data: 'action',
                    name: 'action',
                    orderable: false,
                    searchable: false
                },
            ],
            language: {
                paginate: {
                    next: '<i class="fa fa-angle-right"></i>', // or '→'
                    previous: '<i class="fa fa-angle-left"></i>' // or '←' 
                }
            },
            "aaSorting": [],
            "pageLength": 25
        });

        $('.dataTables_wrapper').addClass('position-relative');
        $('.dataTables_info').addClass('position-absolute');
        $('.dataTables_filter label input').addClass('form-control form-control-solid w-250px ps-14');
        $('.dataTables_filter').addClass('position-absolute end-0 top-0');
        $('.dataTables_length label select').addClass('form-control form-control-solid');

        document.querySelector('#institution_datable_search').addEventListener("keyup", function(e) {
                dtTable.draw();
            }),

            $('#search-form').on('submit', function(e) {
                dtTable.draw();
                e.preventDefault();
            });
        $('#search-form').on('reset', function(e) {
            $('select[name=filter_status]').val(0).change();

            dtTable.draw();
            e.preventDefault();
        });

        function institutionChangeStatus(id, status) {

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
                        url: "{{ route('institutions.change.status') }}",
                        type: 'POST',
                        data: {
                            id: id,
                            status: status
                        },
                        success: function(res) {
                            dtTable.ajax.reload();
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

        $('#kt_common_add_form').on('hidden.bs.modal', function() {
            $(this).find('form').trigger('reset');
        })

        function getInstituteModal( id = '') {

            $.ajaxSetup({
                headers: {
                    'X-CSRF-TOKEN': $('meta[name="csrf-token"]').attr('content')
                }
            });
            $.ajax({
                url: "{{ route('institutions.add_edit') }}",
                type: 'POST',
                data: {
                    id: id,
                },
                success: function(res) {
                    $('#kt_dynamic_app').modal('show');
                    $('#kt_dynamic_app').html(res);
                }
            })

        }

        function deleteInstitution(id) {
            Swal.fire({
                text: "Are you sure you would like to delete record?",
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
                        url: "{{ route('institutions.delete') }}",
                        type: 'POST',
                        data: {
                            id: id,
                        },
                        success: function(res) {
                            dtTable.ajax.reload();
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
       
    </script>
@endsection
