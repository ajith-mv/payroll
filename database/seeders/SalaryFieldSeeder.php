<?php

namespace Database\Seeders;

use App\Models\PayrollManagement\SalaryField;
use Illuminate\Database\Console\Seeds\WithoutModelEvents;
use Illuminate\Database\Seeder;

class SalaryFieldSeeder extends Seeder
{
    
    public function run()
    {
        // $ins['academic_id'] = academicYearId();
        // $ins['name'] = 'Employee Provident Fund';
        // $ins['short_name'] = 'EPF';
        // $ins['description'] = 'Employee Provident Fund';
        // $ins['status'] = 'active';
        // $ins['salary_head_id'] = 2;
        // $ins['entry_type'] = 'calculation';
        // $ins['is_static'] = 'no';
        // $ins['order_in_salary_slip'] = 1;
        // $ins['added_by'] = 1;

        // SalaryField::updateOrcreate(['short_name' => 'EPF'], $ins);

        $ins = [];
        $ins['academic_id'] = academicYearId();
        $ins['name'] = 'Employees\' State Insurance';
        $ins['short_name'] = 'ESI';
        $ins['description'] = 'Employees\' State Insurance';
        $ins['status'] = 'active';
        $ins['salary_head_id'] = 2;
        $ins['order_in_salary_slip'] = 2;
        $ins['entry_type'] = 'inbuilt_calculation';
        $ins['is_static'] = 'yes';
        $ins['added_by'] = 1;

        SalaryField::updateOrcreate(['short_name' => 'ESI', 'salary_head_id' => 2], $ins);

        $ins = [];
        $ins['academic_id'] = academicYearId();
        $ins['name'] = 'Bank Loan';
        $ins['short_name'] = 'Bank Loan';
        $ins['description'] = 'Bank Loan';
        $ins['status'] = 'active';
        $ins['salary_head_id'] = 2;
        $ins['order_in_salary_slip'] = 3;
        $ins['entry_type'] = 'inbuilt_calculation';
        $ins['is_static'] = 'yes';
        $ins['added_by'] = 1;

        SalaryField::updateOrcreate(['short_name' => 'Bank Loan', 'salary_head_id' => 2], $ins);

        $ins = [];
        $ins['academic_id'] = academicYearId();
        $ins['name'] = 'Life Insurance Corporation';
        $ins['short_name'] = 'LIC';
        $ins['description'] = 'Life Insurance Corporation';
        $ins['status'] = 'active';
        $ins['salary_head_id'] = 2;
        $ins['order_in_salary_slip'] = 4;
        $ins['entry_type'] = 'inbuilt_calculation';
        $ins['is_static'] = 'yes';
        $ins['added_by'] = 1;

        SalaryField::updateOrcreate(['short_name' => 'LIC', 'salary_head_id' => 2], $ins);

        $ins = [];
        $ins['academic_id'] = academicYearId();
        $ins['name'] = 'Professional Tax';
        $ins['short_name'] = 'PT';
        $ins['description'] = 'Professional Tax';
        $ins['status'] = 'active';
        $ins['salary_head_id'] = 2;
        $ins['order_in_salary_slip'] = 5;
        $ins['entry_type'] = 'inbuilt_calculation';
        $ins['is_static'] = 'yes';
        $ins['added_by'] = 1;

        SalaryField::updateOrcreate(['short_name' => 'PT', 'salary_head_id' => 2], $ins);

        $ins = [];
        $ins['academic_id'] = academicYearId();
        $ins['name'] = 'Income Tax';
        $ins['short_name'] = 'IT';
        $ins['description'] = 'Income Tax';
        $ins['status'] = 'active';
        $ins['salary_head_id'] = 2;
        $ins['order_in_salary_slip'] = 6;
        $ins['entry_type'] = 'inbuilt_calculation';
        $ins['is_static'] = 'yes';
        $ins['added_by'] = 1;

        SalaryField::updateOrcreate(['short_name' => 'IT', 'salary_head_id' => 2], $ins);

        $ins = [];
        $ins['academic_id'] = academicYearId();
        $ins['name'] = 'Arrears';
        $ins['short_name'] = 'Arrears';
        $ins['description'] = 'Arrears';
        $ins['status'] = 'active';
        $ins['salary_head_id'] = 2;
        $ins['order_in_salary_slip'] = 7;
        $ins['entry_type'] = 'inbuilt_calculation';
        $ins['is_static'] = 'yes';
        $ins['added_by'] = 1;

        SalaryField::updateOrcreate(['short_name' => 'Arrears', 'salary_head_id' => 2], $ins);

        $ins = [];
        $ins['academic_id'] = academicYearId();
        $ins['name'] = 'Contributions';
        $ins['short_name'] = 'Contributions';
        $ins['description'] = 'Contributions';
        $ins['status'] = 'active';
        $ins['salary_head_id'] = 2;
        $ins['order_in_salary_slip'] = 8;
        $ins['entry_type'] = 'inbuilt_calculation';
        $ins['is_static'] = 'yes';
        $ins['added_by'] = 1;

        SalaryField::updateOrcreate(['short_name' => 'Contributions', 'salary_head_id' => 2], $ins);

        $ins = [];
        $ins['academic_id'] = academicYearId();
        $ins['name'] = 'Others';
        $ins['short_name'] = 'Others';
        $ins['description'] = 'Others';
        $ins['status'] = 'active';
        $ins['salary_head_id'] = 2;
        $ins['order_in_salary_slip'] = 9;
        $ins['entry_type'] = 'inbuilt_calculation';
        $ins['is_static'] = 'yes';
        $ins['added_by'] = 1;

        SalaryField::updateOrcreate(['short_name' => 'Others', 'salary_head_id' => 2], $ins);

        $ins = [];
        $ins['academic_id'] = academicYearId();
        $ins['name'] = 'Bonus';
        $ins['short_name'] = 'Bonus';
        $ins['description'] = 'Bonus';
        $ins['status'] = 'active';
        $ins['salary_head_id'] = 1;
        $ins['order_in_salary_slip'] = 10;
        $ins['entry_type'] = 'inbuilt_calculation';
        $ins['is_static'] = 'yes';
        $ins['added_by'] = 1;

        SalaryField::updateOrcreate(['short_name' => 'Bonus', 'salary_head_id' => 1], $ins);

        $ins = [];
        $ins['academic_id'] = academicYearId();
        $ins['name'] = 'Others';
        $ins['short_name'] = 'Others';
        $ins['description'] = 'Others';
        $ins['status'] = 'active';
        $ins['salary_head_id'] = 1;
        $ins['order_in_salary_slip'] = 11;
        $ins['entry_type'] = 'inbuilt_calculation';
        $ins['is_static'] = 'yes';
        $ins['added_by'] = 1;

        SalaryField::updateOrcreate(['short_name' => 'Others', 'salary_head_id' => 1], $ins);

        $ins = [];
        $ins['academic_id'] = academicYearId();
        $ins['name'] = 'Arrears';
        $ins['short_name'] = 'Arrears';
        $ins['description'] = 'Arrears';
        $ins['status'] = 'active';
        $ins['salary_head_id'] = 1;
        $ins['order_in_salary_slip'] = 12;
        $ins['entry_type'] = 'inbuilt_calculation';
        $ins['is_static'] = 'yes';
        $ins['added_by'] = 1;

        SalaryField::updateOrcreate(['short_name' => 'Arrears', 'salary_head_id' => 1], $ins);



    }
}
