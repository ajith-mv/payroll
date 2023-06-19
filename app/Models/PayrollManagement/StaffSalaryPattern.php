<?php

namespace App\Models\PayrollManagement;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use OwenIt\Auditing\Contracts\Auditable;

class StaffSalaryPattern extends Model implements Auditable
{
    use HasFactory;
    use \OwenIt\Auditing\Auditable;

    protected $fillable = [
        'staff_id',
        'salary_no',
        'total_earnings',
        'total_deductions',
        'gross_salary',
        'net_salary',
        'salary_month',
        'salary_year',
        'effective_from',
        'is_salary_processed',
        'salary_approved_by',
        'salary_processed_on',
        'status',
        'payout_month',
        'remarks',
        'employee_remarks',
        'verification_status'
    ];

}