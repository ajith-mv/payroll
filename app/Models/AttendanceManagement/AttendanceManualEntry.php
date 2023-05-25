<?php

namespace App\Models\AttendanceManagement;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class AttendanceManualEntry extends Model
{
    use HasFactory,SoftDeletes;
    protected $fillable = [
        'academic_id',
        'employment_id',
        'attendance_date',
        'reporting_manager',
        'attendance_status',
        'absent_status',
        'reason',
        'sort_order',
        'reason',
        'status',
        'from_time',
        'to_time'
    ];
}