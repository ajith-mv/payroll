<?php

namespace App\Exports;

use App\Models\Master\Nationality;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithHeadings;

class NationalityExport implements FromCollection,WithHeadings
{
    /**
    * @return \Illuminate\Support\Collection
    */
    public function collection()
    {
        return Nationality::select('name','status','created_at')->get();
    }
    public function headings(): array
    {
        return [
        'Name',
        'Status',
        'Created At',
        ]; 
    }
}
