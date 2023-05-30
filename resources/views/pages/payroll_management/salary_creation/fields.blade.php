<hr>
<div class="row mt-3">
    <div class="col-sm-12 m-auto">
        
        <div class="accordion" id="accordionPanelsStayOpenExample">
            <div class="row">
            @isset($salary_heads)
                @foreach ($salary_heads as $item)
                <div class="col-sm-6">
                    <div class="accordion-item">
                        <h2 class="accordion-header" id="panelsStayOpen-headingOne">
                            <button class="accordion-button" type="button" data-bs-target="#panelsStayOpen-collapseOne"
                                aria-expanded="true" aria-controls="panelsStayOpen-collapseOne">
                                {{ $item->name }}
                            </button>
                        </h2>
                        <div id="panelsStayOpen-collapseOne" class="accordion-collapse collapse show"
                            aria-labelledby="panelsStayOpen-headingOne">
                            <div class="accordion-body">
                                <div class="list-group">
                                    @if (isset($item->fields) && !empty($item->fields))
                                        @foreach ($item->fields as $item_fields)
                                        @if( isset( $salary_info ) && !empty( $salary_info) )
                                            @php
                                                $old_data = getSalarySelectedFields($salary_info->staff_id, $salary_info->id, $item_fields->id )
                                            @endphp
                                        @endif
                                            <label class="list-group-item p-3 d-flex justify-content-between">
                                                <input class="form-check-input me-1" type="checkbox"
                                                    data-id="{{ str_replace(' ', '_', $item_fields->short_name) }}"
                                                    onchange="getInputValue(this)" value="" @if( isset( $old_data ) && !empty( $old_data ) ) checked @endif>
                                                <span class="px-3 w-50"> {{ $item_fields->name }}
                                                    ({{ $item_fields->short_name }})
                                                    @if (isset($item_fields->field_items) && count($item_fields->field_items) > 0)
                                                        [
                                                        @foreach ($item_fields->field_items as $sfield_items)
                                                            {{ $sfield_items->field_name }}*{{ $sfield_items->percentage }}%
                                                        @endforeach
                                                        ]
                                                    @endif

                                                </span>
                                                <input type="text" name="amount_{{ $item_fields->id }}"
                                                    onkeyup="getNetSalary(this.value)"
                                                    id="{{ str_replace(' ', '_', $item_fields->short_name) }}_input"
                                                    value="{{ $old_data->amount ?? '' }}"
                                                    @if ($item_fields->no_of_numerals) maxlength="{{ $item_fields->no_of_numerals }}" @endif
                                                    class="border border-2 float-end text-end price @if ($item->id == '1') add_input @else minus_input @endif @if (isset($item_fields->field_items) && count($item_fields->field_items) > 0) automatic_calculation @endif"
                                                    data-id="" @if( isset( $old_data ) && !empty( $old_data ) ) @else disabled @endif >
                                            </label>
                                        @endforeach
                                    @endif
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                   
                @endforeach
                <div class="col-sm-12">

                    <h2 class="accordion-header netsalary" id="panelsStayOpen-headingOne">
                        <input type="hidden" name="net_salary" id="net_salary" value="">
                        Net Salary
                        <span class="float-end">₹ <span id="net_salary_text">{{ $salary_info->net_salary ?? '0.00' }}</span></span>
                    </h2>
                </div>
            @endisset


        </div>
        <div class="form-group mt-5 text-end">
            <button class="btn btn-primary btn-sm" type="submit"> Submit & Lock </button>
            <a class="btn btn-dark btn-sm" href="{{ route('salary.creation') }}"> Cancel </a>
        </div>

    </div>
</div>
