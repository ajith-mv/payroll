<form action="" class="" id="dynamic_form">
    <input type="hidden" name="id" value="{{ $info->id ?? '' }}">
    <input type="hidden" name="form_type" id="form_type" value="{{ $from ?? '' }}">
    <div class="fv-row form-group mb-10">
        <label class="form-label required" for="">
            Designation
        </label>
        <div >
            <input type="text" class="form-control" name="designation" value="{{ $info->name ?? ''}}" id="designation" required >
        </div>
    </div>
    <div class="fv-row form-group mb-10">
        <label class="form-label" for="">
            Can Assignable to Reporting Manager
        </label>
        <div >
            <input type="radio" id="yes" class="form-check-input" value="1" name="can_assign_report_manager" @if(isset($info->can_assign_report_manager) && $info->can_assign_report_manager == 'yes') checked  @endif >
            <label class="pe-3" for="yes">Yes</label>
            <input type="radio" id="no" class="form-check-input" value="0" name="can_assign_report_manager" @if(isset($info->can_assign_report_manager) && $info->can_assign_report_manager == 'no') checked @elseif(!isset($info->can_assign_report_manager) || (isset($info->can_assign_report_manager) && empty( $info->can_assign_report_manager ) ) ) checked @endif >
            <label for="no">No</label>
        </div>
    </div>
    @if(isset($from) && !empty($from))
    <div class="fv-row form-group mb-10">
        <label class="form-label" for="">
            Status
        </label>
        <div >
            <input type="radio" id="active" class="form-check-input" value="1" name="status" @if(isset($info->status) && $info->status == 'active') checked @elseif(!isset($info->status)) checked @endif >
            <label class="pe-3" for="active">Active</label>
            <input type="radio" id="inactive" class="form-check-input" value="0" name="status" @if(isset($info->status) && $info->status != 'active') checked  @endif >
            <label for="inactive">Inactive</label>
        </div>
    </div>
    @endif
    <div class="form-group mb-10 text-end">
        <button type="button" class="btn btn-light-primary" data-bs-dismiss="modal"> Cancel </button>
        <button type="button" class="btn btn-primary" id="form-submit-btn"> 
            <span class="indicator-label">
                Submit
            </span>
            <span class="indicator-progress">
                Please wait... <span class="spinner-border spinner-border-sm align-middle ms-2"></span>
            </span>
        </button>
    </div>
</form>

<script>
    var from = '{{ $from ?? '' }}';

var KTAppEcommerceSaveDesignation = function () {

    const handleSubmit = () => {
        // Define variables
        let validator;
        // Get elements
        const form = document.getElementById('dynamic_form');
        const submitButton = document.getElementById('form-submit-btn');

        validator = FormValidation.formValidation(
            form,
            {
                fields: {
                    'designation': {
						validators: {
							notEmpty: {
								message: 'Designation is required'
							},
						}
					},
                },
                plugins: {
                    trigger: new FormValidation.plugins.Trigger(),
                    bootstrap: new FormValidation.plugins.Bootstrap5({
                        rowSelector: '.fv-row',
                        eleInvalidClass: '',
                        eleValidClass: ''
                    })
                }
            }
        );

        // Handle submit button
        submitButton.addEventListener('click', e => {
            e.preventDefault();
            // Validate form before submit
            if (validator) {
                validator.validate().then(function (status) {

                    if (status == 'Valid') {

                        var forms = $('#dynamic_form')[0];
                        var formData = new FormData(forms);
                        $.ajax({
                            url:"{{ route('save.designation') }}",
                            type:"POST",
                            data: formData,
                            processData: false,
                            contentType: false,
                            success: function(res) {
                                // Disable submit button whilst loading
                                submitButton.disabled = false;
                                submitButton.removeAttribute('data-kt-indicator');
                                if( res.error == 1 ) {
                                    if( res.message ) {
                                        res.message.forEach(element => {
                                            toastr.error("Error", element);
                                        });
                                    }
                                } else{
                                    toastr.success("Designation added successfully");
                                    $('#kt_dynamic_app').modal('hide');

                                    if (from) {
                                            dtTable.draw();
                                        } else {
                                    if( res.inserted_data ) {
                                        $('#designation_id').append(`<option value="${res.inserted_data.id}">${res.inserted_data.name}</option>`)
                                        $('#designation_id').val(res.inserted_data.id).trigger('change');
                                    }
                                }
                                }
                            }
                        })

                    } 
                });
            }
        })
    }

    return {
        init: function () {
            handleSubmit();
        }
    };
}();

KTUtil.onDOMContentLoaded(function () {
    KTAppEcommerceSaveDesignation.init();
});
    
    
</script>