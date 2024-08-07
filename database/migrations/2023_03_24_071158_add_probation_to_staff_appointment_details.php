<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

return new class extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::table('staff_appointment_details', function (Blueprint $table) {
            $table->string('probation_period')->nullable()->after('has_probation');
            $table->string('status')->nullable()->after('probation_period');
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::table('staff_appointment_details', function (Blueprint $table) {
            $table->dropColumn('probation_period');
            $table->dropColumn('status');
        });
    }
};
