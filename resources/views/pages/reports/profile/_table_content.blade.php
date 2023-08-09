@if (isset($details) && !empty($details))
    @foreach ($details as $item)
        <tr>
            <td>
                {{ $loop->iteration }}
            </td>
            <td>
                @if (getStaffImage($item->image))
                    <img src="{{ getStaffImage($item->image) }}" alt="" width="100">
                @endif
            </td>
            <td>{{ $item->name ?? '' }}</td>
            <td>{{ $item->emp_code ?? '' }}</td>
            <td>{{ $item->society_emp_code ?? '' }}</td>
            <td> {{ $item->institute_emp_code ?? '' }}</td>
            <td>{{ isset($item->personal->dob) && !empty( $item->personal->dob) ? commonDateFormat($item->personal->dob) : '-'  }}</td>
            <td>{{ ucfirst($item->personal->gender ?? '') }}</td>
            <td>{{ $item->position->designation->name ?? '-' }}</td>
            {{-- <td>Place Of Work</td> --}}
            <td>{{ $item->personal->motherTongue->name ?? '-' }}</td>
            <td>{{ $item->personal->mobile_no1 ?? '' }}</td>
            <td>{{ $item->personal->whatsapp_no ?? '' }}</td>
            <td>{{ $item->personal->emergency_no ?? '' }}</td>
            <td>{{ $item->personal->birthPlace->name ?? '-' }}</td>
            <td>{{ $item->personal->nationality->name ?? '-' }}</td>
            <td>{{ $item->personal->religion->name ?? '-' }}</td>
            <td>{{ $item->personal->caste->name ?? '-' }}</td>
            <td>{{ $item->personal->community->name ?? '-' }}</td>
            <td>{{ $item->personal->contact_address ?? '' }}</td>
            <td>{{ $item->personal->permanent_address ?? '' }}</td>
            <td>{{ $item->personal->marital_status ?? '' }}</td>
            <td>{{ isset($item->personal->marriage_date) && !empty($item->personal->marriage_date ) ? commonDateFormat($item->personal->marriage_date) : '' }}</td>
            {{-- <td>Adhaar</td>
            <td> Pan Card </td>
            <td> Ration Card </td>
            <td> Voter ID </td> --}}
        </tr>
    @endforeach
@else
@endif