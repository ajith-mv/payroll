<?php

namespace App\Imports;

use App\Models\User;
use Carbon\Carbon;
use Illuminate\Support\Collection;
use Maatwebsite\Excel\Concerns\ToCollection;
use App\Models\PayrollManagement\StaffSalaryField;
use App\Models\PayrollManagement\StaffSalaryPattern;
use App\Models\PayrollManagement\StaffSalaryPatternField;
use App\Models\PayrollManagement\StaffSalaryPatternFieldHistory;
use App\Models\PayrollManagement\StaffSalaryPatternHistory;
use Maatwebsite\Excel\Concerns\WithHeadingRow;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use App\Models\PayrollManagement\SalaryField;
use App\Models\PayrollManagement\Payroll;
use App\Models\PayrollManagement\PayrollPermission;
use App\Http\Controllers\PayrollManagement\OverviewController;
use App\Models\Staff\StaffTaxSeperation;
use App\Models\PayrollManagement\ItStaffStatement;
use App\Repositories\TaxCalculationRepository;
use App\Repositories\PayrollChecklistRepository;
use App\Models\Staff\StaffBankLoan;
use App\Models\Staff\StaffInsurance;
use App\Models\Staff\StaffInsuranceEmi;
use App\Models\Staff\StaffLoanEmi;
use App\Models\Staff\StaffSalaryPreDeduction;
use App\Models\Staff\StaffSalaryPreEarning;
use DB;
use Illuminate\Support\Str;
use App\Models\PayrollManagement\StaffSalary;
class PayrollImport implements ToCollection,WithHeadingRow
{
    public $collection;

    public function collection(Collection $rows)
    {
      
        set_time_limit(0);
        ini_set('memory_limit', '-1');
       $dateTime =  date('Y-m-d', strtotime('2024-12-01'));
       $date= date('Y-m-d', strtotime('2024-12-01'));
      $payroll_date = $date;
      $from_date = date('Y-m-01', strtotime($payroll_date));
      $to_date = date('Y-m-t', strtotime($payroll_date));

      $ins['from_date'] = $from_date;
      $ins['to_date'] = $to_date;
      $ins['name'] = date('F Y', strtotime($from_date));
      $ins['locked'] = 'no';
      $ins['added_by'] = auth()->id();
      $ins['academic_id'] = academicYearId();
      $ins['institute_id'] = session()->get('staff_institute_id')??null;
      $payroll = Payroll::updateOrCreate(['academic_id'=>academicYearId(),'institute_id'=>session()->get('staff_institute_id'),'from_date'=>$from_date],$ins);
    
      $ins_roll['academic_id'] = academicYearId();
      $ins_roll['payout_month'] = $from_date;
      $ins_roll['payroll_id'] = $payroll->id;
      $ins_roll['payroll_inputs'] = 'unlock';
      $ins_roll['emp_view_release'] = 'lock';
      $ins_roll['it_statement_view'] = 'lock';
      $ins_roll['payroll'] = 'unlock';
      $academicId = academicYearId();
      PayrollPermission::updateOrCreate(['payroll_id'=>$payroll->id],$ins_roll);
      User::where('institute_id', session()->get('staff_institute_id'))
            ->where('status', 'active')
            ->where('verification_status', 'approved')
            ->where('transfer_status', 'active')
            ->chunk(100, function($incomes) use ($ins,$academicId) {
            
                foreach ($incomes as $income) {
                    $income_data = ItStaffStatement::where('academic_id',academicYearId())->where('staff_id',$income->id)->first();
                    if(empty($income_data)){
                      $taxRepo = new TaxCalculationRepository();
                      $result = $taxRepo->generateStatementForStaff($income->id);
                    }
                    $income_info = ItStaffStatement::where('academic_id',academicYearId())->where('staff_id',$income->id)->first();
                    if (isset($income_info) ){
                        $total_income_tax_payable = $income_info->total_income_tax_payable ?? 0;
                        $tax_amount = $total_income_tax_payable / 4;

                        $ins_tax = [
                            'academic_id' => $academicId,
                            'staff_id' => $income->id,
                            'income_tax_id' => $income_info->id,
                            'april' => 0,
                            'may' => 0,
                            'june' => 0,
                            'july' => 0,
                            'august' => 0,
                            'september' => 0,
                            'october' => 0,
                            'november' => 0,
                            'december' => $tax_amount,
                            'january' => $tax_amount,
                            'february' => $tax_amount,
                            'march' => $tax_amount,
                            'total_tax' => $total_income_tax_payable
                        ];

                        StaffTaxSeperation::updateOrCreate(
                            ['staff_id' => $income->id, 'income_tax_id' => $income_info->id],
                            $ins_tax
                        );
                        $income_info->update(['is_staff_calculation_done' => 'yes']);
                    }
                }
            });
                ini_set("max_execution_time", 0);
                ini_set('memory_limit', '-1');
                $batchSize = 4000;
                $payout_id = $payroll->id;
                $payroll_date = date('Y-m-d', strtotime($date));
                $salary_month = date('F', strtotime($payroll_date));
                $salary_year = date('Y', strtotime($payroll_date));
                $month = date('m', strtotime($payroll_date));
                $month_length = date('t', strtotime($payroll_date));
                $total_net_pay = 0;

                $month_start = date('Y-m-01', strtotime($payroll_date));
                $month_end = date('Y-m-t', strtotime($payroll_date));
                $working_day = date('t', strtotime($payroll_date));
                try {
                DB::transaction(function() use ($date,$payout_id,$payroll_date,$salary_month,$salary_year,$month_length,$total_net_pay,
                $working_day,$month_start,$month_end,$batchSize,$rows,$month
                ) {
              
                    
                    $payCheck = new PayrollChecklistRepository();
                    $payout_data = $payCheck->getToPayEmployee($date);
                if (isset($rows) && count($rows)) {
        
                    StaffSalary::where('payroll_id', $payout_id)->update(['status' => 'inactive']);
                    $ins=[];
                    foreach ($rows as $key => $row) {
                       
                     $staff_info =User::where('institute_emp_code',$row["inst_emp_code"])->first();
                     $earings_field = SalaryField::where('salary_head_id', 1)->where('nature_id',  $staff_info->appointment->employment_nature->id?? 3 )
                        ->get();
                     $deductions_field = SalaryField::where('salary_head_id', 2)
                     ->where(function ($query) {
                         $query->whereNull('nature_id')
                               ->orWhere('nature_id', 3);
                     })
                     ->whereNull('deleted_at')
                     ->orderBy('order_in_salary_slip', 'asc')
                     ->get();
                        if (isset($staff_info) && !empty($staff_info)) {
                            $total_earnings=0;
                            $total_deductions=0;
                            $staff_id = $staff_info->id;
                            $sal['staff_id'] = $staff_id;
                            $sal['payroll_id'] = $payout_id;
                            $sal['salary_month'] = $salary_month;
                            $sal['salary_year'] = $salary_year;
                            $sal['is_salary_processed'] = 'yes';
                            $sal['status'] = 'active';
                            
                            $sal['salary_pattern_id'] = $staff_info->currentSalaryPattern->id ?? 1;
                            $sal['working_days'] = $working_day;
                            $sal['worked_days'] = $staff_info->workedDays->count();
                            $sal['other_description'] = NUll;
                            $sal['salary_date'] = $payroll_date;
                            $sal['salary_no'] =  salaryNo();
                            $sal['is_salary_processed'] = 'yes';
                            $sal['status'] = 'active';
                            $sal['salary_processed_on']= date('Y-m-d H:i:s');
                            $ins[]=$sal;
                            $sallary_f_id=StaffSalary::updateOrCreate(
                              [
                                  'staff_id' => $staff_id,
                                  'payroll_id' => $payout_id,
                              ],
                              $sal
                          );
                        
                            if (isset($earings_field) && !empty($earings_field)) {
                            
                                foreach ($earings_field as $eitem) {
                                    if($eitem->entry_type=="calculation"){
                                        $valuesArray = explode(',', $eitem->field_items->field_name);
                                        $C_amount=0;
                                        if(isset($valuesArray[0]) && !empty($valuesArray[0])){
                                            if(isset($eitem->PrecentageLog) && !empty($eitem->PrecentageLog)){
                                                $percentage=$eitem->PrecentageLog->new_percentage;
                                            }else{
                                                $percentage=$eitem->field_items->percentage;
                                            }
                                           
                                            $C_amount +=($percentage/100)* $row[strtolower($valuesArray[0])];
                                           
                                        }
                                        if(isset($valuesArray[1]) && !empty($valuesArray[1])){
                                            if($valuesArray[1]=="DA"){
                                                $da=($eitem->field_items->percentage/100)* $row[strtolower($valuesArray[0])]; 
                                                $C_amount +=($eitem->field_items->percentage/100)* $da;
                                            }
                                           
                                        }
    if(isset($staff_info->appointment) && isset($staff_info->appointment->employment_nature) && isset($staff_info->appointment->employment_nature->id)&& in_array($staff_info->appointment->employment_nature->id, [4, 1]) &&in_array($eitem->short_name, ['HRA', 'DA']) ){
        $C_amount=0;
        }
                                        $used_fields= [
                                            'percentage' => 0,
                                            'staff_id' => $staff_info->id,
                                            'field_id' => $eitem->id,
                                            'field_name' => $eitem->name,
                                            'reference_type' => 'EARNINGS',
                                            'reference_id' => 1,
                                            'short_name' => $eitem->short_name,
                                            'staff_salary_id'=>$sallary_f_id->id,
                                            'amount' => round($C_amount),
                                        ];
                                        $total_earnings +=round($C_amount);
                                       }else{
                                        switch (strtolower(trim($eitem->short_name))) {
                                        case 'arr':
                                        $amount=StaffSalaryPreEarning::where('staff_id',$staff_info->id)->where('earnings_type','arrear')->whereMonth('salary_month', $month)->whereYear('salary_month', $salary_year)->where('status','active')->sum('amount');
                                        $M_amount = $amount??0;
                                        break;
                                        case 'bonus':
                                        $amount=StaffSalaryPreEarning::where('staff_id',$staff_info->id)->where('earnings_type','bonus')->whereMonth('salary_month', $month)->whereYear('salary_month', $salary_year)->where('status','active')->sum('amount');
                                        $M_amount = $amount??0;
                                        break;
                                        case 'others':
                                        $amount=StaffSalaryPreEarning::where('staff_id',$staff_info->id)->where('earnings_type','other')->whereMonth('salary_month', $month)->whereYear('salary_month', $salary_year)->where('status','active')->sum('amount');
                                        $M_amount = $amount??0;
                                        break;
                                        default:
                                        $M_amount = $row[strtolower($eitem->short_name)] ?? 0;
                                        break;
                                        }
                                        $used_fields= [
                                            'percentage' => 0,
                                            'staff_id' => $staff_info->id,
                                            'field_id' => $eitem->id,
                                            'field_name' => $eitem->name,
                                            'reference_type' => 'EARNINGS',
                                            'reference_id' => 1,
                                            'short_name' => $eitem->short_name,
                                            'staff_salary_id'=>$sallary_f_id->id,
                                            'amount' => round($M_amount),
                                            
                                        ];
                                        $total_earnings +=round($M_amount);
                                       }
                                
                                StaffSalaryField::updateOrCreate($used_fields,['staff_id'=>$staff_id,'staff_salary_id'=>$sallary_f_id->id,'field_id'=> $eitem->id]);
                                    
                                }
                            }
                            if (isset($deductions_field) && !empty($deductions_field)) {
                                foreach ($deductions_field as $sitem) {
                                    if($sitem->entry_type=="calculation"){
                                        $valuesArray = explode(',', $sitem->field_items->field_name);
                                      
                                        $D_amount=0;
                                    if(isset($valuesArray[0]) && !empty($valuesArray[0])){
                                        if(isset($sitem->PrecentageLog) && !empty($sitem->PrecentageLog)){
                                            $percentage=$sitem->PrecentageLog->new_percentage;
                                        }else{
                                            $percentage=$sitem->field_items->percentage;
                                        }
                                        $D_amount +=($percentage/100)* $row[strtolower($valuesArray[0])]; 
                                    }
                                  
                                    if(isset($valuesArray[1]) && !empty($valuesArray[1])){
                                        if($valuesArray[1]=="DA"){
                                            $da_field = SalaryField::with('PrecentageLog','field_items')->where('salary_head_id', 1)->where('nature_id',  $staff_info->appointment->employment_nature->id?? 1 )->where('short_name','DA')->first();                                            $da=($sitem->field_items->percentage/100)* $row[strtolower($valuesArray[0])]; 
                                            if(isset($da_field->PrecentageLog) && !empty($da_field->PrecentageLog)){
                                                $percentage=$da_field->PrecentageLog->new_percentage;
                                            }else{
                                                if(isset($da_field->field_items)){
                                                    $percentage=$da_field->field_items->percentage;
                                                }else{
                                                    $percentage=12;
                                                }
                                              
                                            }
                                            if(isset($da_field) && isset($da_field->field_items)){
                                              $D_amount +=($percentage/100)* $da;
                                           }
                                            
                                        }
                                        
                                    }
                                  
                                      if(isset($staff_info->appointment) && isset($staff_info->appointment->employment_nature) && isset($staff_info->appointment->employment_nature->id)&& in_array($staff_info->appointment->employment_nature->id, [5, 4, 1]) && $sitem->short_name=='EPF'){
                                        $D_amount=0;
                                      }
                                        $tmp= [
                                            'percentage' => 0,
                                            'staff_id' => $staff_info->id,
                                            'field_id' => $sitem->id,
                                            'field_name' => $sitem->name,
                                            'reference_type' => 'DEDUCTIONS',
                                            'reference_id' => 1,
                                            'short_name' => $sitem->short_name,
                                            'staff_salary_id'=>$sallary_f_id->id,
                                            'amount' => round($D_amount),
                                        ];
                                        $total_deductions +=round($D_amount);
                                       }else{
                                        switch (strtolower(trim($sitem->short_name))) {
                                            case 'contri':
                                                $amount=StaffSalaryPreDeduction::where('staff_id',$staff_info->id)->where('deduction_type','contribution')->whereMonth('salary_month', $month)->whereYear('salary_month', $salary_year)->where('status','active')->sum('amount');
                                                $deduct_amount = $amount??0;
                                                break;
                                            case 'other':
                                                $amount=StaffSalaryPreDeduction::where('staff_id',$staff_info->id)->where('deduction_type','other')->whereMonth('salary_month', $month)->whereYear('salary_month', $salary_year)->where('status','active')->sum('amount');
                                                $deduct_amount = $amount??0;
                                                break;
                                            case 'bank loan':
                                                $other_bank_loan_amount = getBankLoansAmount($staff_info->id, $payroll_date,'home_loan');
                                                if (!empty($other_bank_loan_amount['emi'])) {
                                                    $used_loans[] = $other_bank_loan_amount['emi'];
                                                }
                                           
                                                if (!empty($used_loans)) {
                                                    $staff_loan_id = [];
                                                    foreach ($used_loans as $loan_items) {
                                                        if (isset($loan_items['details']) && !empty($loan_items['details'])) {
                                
                                                            $info = StaffLoanEmi::find($loan_items['details']->id);
                                                            $info->status = 'paid';
                                                            $info->save();
                                                            $staff_loan_id[]  = $loan_items['details']->staff_loan_id;
                                                        }
                                                    }
                                                    if (!empty($staff_loan_id)) {
                                                        $staff_loan_id = array_unique($staff_loan_id);
                                                        foreach ($staff_loan_id as $loan_ids) {
                                                            $loan_info = StaffBankLoan::with('paid_emi')->find($loan_ids);
                                                            if ($loan_info && $loan_info->period_of_loans == $loan_info->paid_emi()->count()) {
                                                                $loan_info->status = 'completed';
                                                                $loan_info->save();
                                                            }
                                                        }
                                                    }
                                                }
                                                $deduct_amount = $other_bank_loan_amount['total_amount']??0;
                                                break;
                                                case 'pl':
                                                    $other_personal_loan_amount = getBankLoansAmount($staff_info->id, $payroll_date,'personal_loan');

                                                    if (!empty($other_personal_loan_amount['emi'])) {
                                                        $used_loans[] = $other_personal_loan_amount['emi'];
                                                    }
                                               
                                                    if (!empty($used_loans)) {
                                                        $staff_loan_id = [];
                                                        foreach ($used_loans as $loan_items) {
                                                            if (isset($loan_items['details']) && !empty($loan_items['details'])) {
                                    
                                                                $info = StaffLoanEmi::find($loan_items['details']->id);
                                                                $info->status = 'paid';
                                                                $info->save();
                                                                $staff_loan_id[]  = $loan_items['details']->staff_loan_id;
                                                            }
                                                        }
                                                        if (!empty($staff_loan_id)) {
                                                            $staff_loan_id = array_unique($staff_loan_id);
                                                            foreach ($staff_loan_id as $loan_ids) {
                                                                $loan_info = StaffBankLoan::with('paid_emi')->find($loan_ids);
                                                                if ($loan_info && $loan_info->period_of_loans == $loan_info->paid_emi()->count()) {
                                                                    $loan_info->status = 'completed';
                                                                    $loan_info->save();
                                                                }
                                                            }
                                                        }
                                                    }
                                                    $deduct_amount = $other_personal_loan_amount['total_amount']??0;
                                                   
                                                    break;
                                                    case 'ol':
                                                        $other_loan_amount = getBankLoansAmount($staff_info->id, $payroll_date,'other_loan');
                                                        if (!empty($other_loan_amount['emi'])) {
                                                            $used_loans[] = $other_loan_amount['emi'];
                                                        }
                                                   
                                                        if (!empty($used_loans)) {
                                                            $staff_loan_id = [];
                                                            foreach ($used_loans as $loan_items) {
                                                                if (isset($loan_items['details']) && !empty($loan_items['details'])) {
                                        
                                                                    $info = StaffLoanEmi::find($loan_items['details']->id);
                                                                    $info->status = 'paid';
                                                                    $info->save();
                                                                    $staff_loan_id[]  = $loan_items['details']->staff_loan_id;
                                                                }
                                                            }
                                                            if (!empty($staff_loan_id)) {
                                                                $staff_loan_id = array_unique($staff_loan_id);
                                                                foreach ($staff_loan_id as $loan_ids) {
                                                                    $loan_info = StaffBankLoan::with('paid_emi')->find($loan_ids);
                                                                    if ($loan_info && $loan_info->period_of_loans == $loan_info->paid_emi()->count()) {
                                                                        $loan_info->status = 'completed';
                                                                        $loan_info->save();
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        $deduct_amount = $other_loan_amount['total_amount']??0;
                                                        break;
                                                case 'lic':
                                            
                                                $other_insurance_amount = getInsuranceAmount($staff_info->id, $payroll_date);
                                             
                                                if (!empty($other_insurance_amount['emi'])) {
                                                    $used_insurance[] = $other_insurance_amount['emi'];
                                                }
                                                if (!empty($used_insurance)) {
                                                    $staff_ins_id = [];
                                                    foreach ($used_insurance as $loan_items) {
                                                        if (isset($loan_items['details']) && !empty($loan_items['details'])) {
                                
                                                            $info = StaffInsuranceEmi::find($loan_items['details']->id);
                                                            $info->status = 'paid';
                                                            $info->save();
                                
                                                            $staff_ins_id[]  = $loan_items['details']->staff_ins_id;
                                                        }
                                                    }
                                                    if (!empty($staff_ins_id)) {
                                                        $staff_ins_id = array_unique($staff_ins_id);
                                                        foreach ($staff_ins_id as $loan_ids) {
                                                            $loan_info = StaffInsurance::with('paid_emi')->find($loan_ids);
                                                            if ($loan_info && $loan_info->period_of_loans == $loan_info->paid_emi()->count()) {
                                                                $loan_info->status = 'completed';
                                                                $loan_info->save();
                                                            }
                                                        }
                                                    }
                                                }
                                                $deduct_amount =  $other_insurance_amount['total_amount'] ??0;
                                                break;
                                            default:
                                                $deduct_amount = $row[strtolower($sitem->short_name)] ?? 0;
                                                break;
                                        }
                                    $tmp= [
                                        'percentage' => 0,
                                        'staff_id' => $staff_info->id,
                                        'field_id' => $sitem->id,
                                        'field_name' => $sitem->name,
                                        'reference_type' => 'DEDUCTIONS',
                                        'reference_id' => 2,
                                        'short_name' => $sitem->short_name,
                                        'staff_salary_id'=>$sallary_f_id->id,
                                        'amount'=> round($deduct_amount)
                                    ];
                                    $total_deductions +=round($deduct_amount);
                                }
                                    StaffSalaryField::updateOrCreate($tmp,['staff_id'=>$staff_id,'staff_salary_id'=>$sallary_f_id->id,'field_id'=> $sitem->id]);
                                }
                            }
                            
                            $sallary_f_id->total_earnings = round($total_earnings);
                            $sallary_f_id->total_deductions = round($total_deductions);
                            $sallary_f_id->gross_salary = round($total_earnings);
                            $sallary_f_id->net_salary= round($total_earnings) - round($total_deductions);
                            $sallary_f_id->update();
                          
                        }
                      }
                    }
                });
                DB::commit();
                }
              
                catch (\Throwable $e) {
                    DB::rollBack();
                  dd($e);
                }
                return true;
            
            

    }
 
  
}
