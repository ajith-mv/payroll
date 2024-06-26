<?php

namespace App\Models\Staff;

use App\Models\Master\RelationshipType;
use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use OwenIt\Auditing\Contracts\Auditable;

class StaffNominee extends Model implements Auditable
{
    use HasFactory;
    use \OwenIt\Auditing\Auditable;
    
    protected $table = 'staff_nominees';
    protected $fillable = [
        'staff_id',
        'academic_id',
        'nominee_id',
        'relationship_type_id',
        'dob',
        'gender',
        'age',
        'minor_name',
        'share',
        'minor_address',
        'minor_contact_no',
        'guardian_name'
    ];

    public function nominee()
    {
        return $this->hasOne(StaffFamilyMember::class, 'id', 'nominee_id');
    }

    public function relationship()
    {
        return $this->hasOne(RelationshipType::class, 'id', 'relationship_type_id');
    }
}
