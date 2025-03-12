<?php

namespace App\Http\Controllers\Admin;

use App\Http\Controllers\Controller;
use App\Models\CertificateTemplate;
use Illuminate\Http\Request;
use App\Models\Certificate;
use App\Models\Student;
use App\Models\StudentEnroll;
use App\Models\StudentAttendance;
use App\Models\Semester;
use App\Models\Subject;
use App\Models\Program;
use App\Models\Batch; 
use App\Models\Grade;
use Toastr;
use PhpOffice\PhpWord\TemplateProcessor;
use Illuminate\Support\Facades\Storage;
use Illuminate\Support\Facades\DB;
use Carbon\Carbon;
class EnrollmentCertificateController extends Controller
{
    /**
     * Create a new controller instance.
     *
     * @return void
     */
    public function __construct()
    {
        // Module Data
        $this->title = trans_choice('module_certificate', 1);
        $this->route = 'admin.enrollment_certificate';
        $this->view = 'admin.enrollment_certificate';
        $this->path = 'enrollment_certificate';
        $this->access = 'enrollment_certificate';


        $this->middleware('permission:'.$this->access.'-view|'.$this->access.'-create|'.$this->access.'-edit|'.$this->access.'-print|'.$this->access.'-download', ['only' => ['index','show']]);
        $this->middleware('permission:'.$this->access.'-create', ['only' => ['create','store']]);
        $this->middleware('permission:'.$this->access.'-edit', ['only' => ['edit','update']]);
        $this->middleware('permission:'.$this->access.'-print', ['only' => ['print']]);
        $this->middleware('permission:'.$this->access.'-download', ['only' => ['download']]);
    }

    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function index(Request $request)
    {
        //
        $data['title'] = $this->title;
        $data['route'] = $this->route;
        $data['view'] = $this->view;
        $data['path'] = $this->path;
        $data['access'] = $this->access;


        if(!empty($request->batch) || $request->batch != null){
            $data['selected_batch'] = $batch = $request->batch;
        }
        else{
            $data['selected_batch'] = '0';
        }

        if(!empty($request->program) || $request->program != null){
            $data['selected_program'] = $program = $request->program;
        }
        else{
            $data['selected_program'] = '0';
        }

        if(!empty($request->student_id) || $request->student_id != null){
            $data['selected_student_id'] = $student_id = $request->student_id;
        }
        else{
            $data['selected_student_id'] = null;
        }

        if(!empty($request->template) || $request->template != null){
            $data['selected_template'] = $template = $request->template;
        }
        else{
            $data['selected_template'] = '0';
        }


        $data['batchs'] = Batch::where('status', '1')
                        ->orderBy('id', 'desc')->get();
        $data['programs'] = Program::where('status', '1')
                        ->orderBy('title', 'asc')->get();
        $data['grades'] = Grade::where('status', '1')
                        ->orderBy('point', 'asc')->get();
        $data['templates'] = CertificateTemplate::where('status', '1')
                        ->orderBy('title', 'asc')->get();
        $data['students'] = Student::whereHas('currentEnroll')->where('status', '1')->orderBy('student_id', 'asc')->get();

        if(!empty($request->template)){

            $data['certificate_template'] = CertificateTemplate::where('id', $template)->first();
        }


        // Filter Student
        if(!empty($request->batch) || !empty($request->program) || !empty($request->student_id) || !empty($request->template)){

            $students = Student::where('status', '!=', '0');

            if(!empty($request->batch) && $request->batch != '0'){
                $students->where('batch_id', $batch);
            }
            if(!empty($request->program) && $request->program != '0'){
                $students->where('program_id', $program);
            }
            if(!empty($request->student_id)){
                $students->where('student_id', 'LIKE', '%'.$student_id.'%');
            }

            $data['rows'] = $students->orderBy('student_id', 'asc')->get();
        }


        // Certificate List
        if(!empty($request->batch) || !empty($request->program) || !empty($request->student_id) || !empty($request->template)){
        
            $certificate = Certificate::where('status', '!=', '0');
            if(!empty($request->template) && $request->template != '0'){
                $certificate->where('template_id', $template);
            }
            if(!empty($request->batch) && $request->batch != '0'){
                $certificate->with('student')->whereHas('student', function ($query) use ($batch){
                    $query->where('batch_id', $batch);
                });
            }
            if(!empty($request->program) && $request->program != '0'){
                $certificate->with('student')->whereHas('student', function ($query) use ($program){
                    $query->where('program_id', $program);
                });
            }
            if(!empty($request->student_id) && $request->student_id != '0'){
                $certificate->with('student')->whereHas('student', function ($query) use ($student_id){
                    $query->where('student_id', 'LIKE', '%'.$student_id.'%');
                });
            }
            $data['certificates'] = $certificate->orderBy('id', 'desc')->get();
        }


        return view($this->view.'.index', $data);
    }

    /**
     * Store a newly created resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @return \Illuminate\Http\Response
     */
    public function store(Request $request)
    {
        // Field Validation
        $request->validate([
            'student_id' => 'required',
            'template_id' => 'required',
            'date' => 'required|date',
            'starting_year' => 'required|numeric',
            'ending_year' => 'required|numeric',
            'credits' => 'required|numeric',
            'point' => 'required|numeric',
        ]);


        $row = Student::where('id', $request->student_id)->first();
        $grades = Grade::where('status', '1')
                        ->orderBy('point', 'asc')->get();

        // Result Calculation
        $total_credits = 0;
        $total_cgpa = 0;
        $starting_year = '0000';
        $ending_year = '0000';

        foreach( $row->studentEnrolls as $key => $item ){
           


            if(isset($item->subjectMarks)){
            foreach($item->subjectMarks as $mark){

                $marks_per = round($mark->total_marks);

                foreach($grades as $grade){
                    if($marks_per >= $grade->min_mark && $marks_per <= $grade->max_mark){
                        if($grade->point > 0){
                        $total_cgpa = $total_cgpa + ($grade->point * $mark->subject->credit_hour);
                        $total_credits = $total_credits + $mark->subject->credit_hour;
                        }
                    break;
                    }
                }
            }}
        }

        $original_credits = $total_credits;
        if($total_credits <= 0){
            $total_credits = 1;
        }
        $com_gpa = $total_cgpa / $total_credits;


        // Insert Data
        $certificate = new Certificate;
        $certificate->template_id = $request->template_id;
        $certificate->student_id = $request->student_id;
        $certificate->date = $request->date;
        $certificate->starting_year = $starting_year;
        $certificate->ending_year = $ending_year;
        $certificate->credits = $original_credits;
        $certificate->point = number_format((float)$com_gpa, 2, '.', '');
        $certificate->status = '1';
        $certificate->save();

        // Set SL No
        $certificate->serial_no = (intval($certificate->id) + intval(100000));
        $certificate->barcode = (intval($certificate->id) + intval(100000));
        $certificate->save();


        Toastr::success(__('msg_updated_successfully'), __('msg_success'));

        return redirect()->back();
    }

    /**
     * Update the specified resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @param  int  $id
     * @return \Illuminate\Http\Response
     */
    public function update(Request $request, Certificate $certificate)
    {
        // Field Validation
        $request->validate([
            'student_id' => 'required',
            'date' => 'required|date',
            'starting_year' => 'required|numeric',
            'ending_year' => 'required|numeric',
            'credits' => 'required|numeric',
            'point' => 'required|numeric',
        ]);


        $row = Student::where('id', $request->student_id)->first();
        $grades = Grade::where('status', '1')->orderBy('point', 'asc')->get();

        // Result Calculation
        $total_credits = 0;
        $total_cgpa = 0;
        $starting_year = '0000';
        $ending_year = '0000';

        foreach( $row->studentEnrolls as $key => $item ){
            if($key == 0){
            $starting_year = $item->session->start_date;
            }
            $ending_year = $item->session->end_date;
            

            if(isset($item->subjectMarks)){
            foreach($item->subjectMarks as $mark){

                $marks_per = round($mark->total_marks);

                foreach($grades as $grade){
                    if($marks_per >= $grade->min_mark && $marks_per <= $grade->max_mark){
                        if($grade->point > 0){
                        $total_cgpa = $total_cgpa + ($grade->point * $mark->subject->credit_hour);
                        $total_credits = $total_credits + $mark->subject->credit_hour;
                        }
                    break;
                    }
                }
            }}
        }

        $original_credits = $total_credits;
        if($total_credits <= 0){
            $total_credits = 1;
        }
        $com_gpa = $total_cgpa / $total_credits;


        // Update Data
        $certificate->date = $request->date;
        $certificate->starting_year = $starting_year;
        $certificate->ending_year = $ending_year;
        $certificate->credits = $original_credits;
        $certificate->point = number_format((float)$com_gpa, 2, '.', '');
        $certificate->save();


        Toastr::success(__('msg_updated_successfully'), __('msg_success'));

        return redirect()->back();
    }

    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function print($id)
    {
        //
        $data['title'] = $this->title;
        $data['route'] = $this->route;
        $data['view'] = $this->view;
        $data['path'] = $this->path;

        // View
        $data['certificate'] = Certificate::where('status', '1')->findOrFail($id);
        $data['grades'] = Grade::where('status', '1')->orderBy('point', 'asc')->get();

        return view($this->view.'.print', $data);
    }

    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function download($id)
    {
        //
        $data['title'] = $this->title;
        $data['route'] = $this->route;
        $data['view'] = $this->view;
        $data['path'] = $this->path;

        // View
        $data['certificate'] = Certificate::where('status', '1')->findOrFail($id);
        $data['grades'] = Grade::where('status', '1')->orderBy('point', 'asc')->get();

        return view($this->view.'.download', $data);
    }
    
    
    
    // public function generateCertificate($id)
    // {
   
    //     // Student ka data fetch karein
    //     $student = Student::findOrFail($id);

    //             // dd($student);  
    //           $semesterName = DB::table('student_enrolls')
    //         ->join('semesters', 'student_enrolls.semester_id', '=', 'semesters.id')
    //         ->join('subjects', 'student_enrolls.subject_id', '=', 'subjects.id')
    //         ->join('programs', 'student_enrolls.program_id', '=', 'programs.id')
    //         ->leftJoin('student_attendances', 'student_attendances.student_id', '=', 'student_enrolls.student_id')
    //         ->where('student_enrolls.student_id', $student->id)
    //         ->select(
    //             'semesters.title as semester_name',
    //             'semesters.start_date as startDate',
    //             'semesters.end_date as endDate',
    //             'subjects.credit_hour as subjectHour',
    //             'programs.title as programTitle',
    //             DB::raw('SUM(student_attendances.present_hours) as totalPresentHours'), // Total present hours
    //             DB::raw('CONCAT(ROUND((SUM(student_attendances.present_hours) * 100) / NULLIF(subjects.credit_hour, 0), 2), " %") as attendancePercentage') // Attendance % with %
    //         )
    //         ->groupBy('semesters.id', 'semesters.title', 'semesters.start_date', 'semesters.end_date', 'subjects.credit_hour', 'programs.title')
    //         ->first();

    //         $subjectIds = DB::table('student_enrolls')
    //             ->where('student_id', $student->id)
    //             ->pluck('subject_id')
    //             ->first();
            
    //         $subjectIdArray = explode(',', $subjectIds);
            
    //      $subjectHours = DB::table('subjects')
    //                     ->whereIn('id', $subjectIdArray)
    //                     ->pluck('credit_hour')
    //                     ->toArray();
             


    //      $issuanceNumber = now()->format('YmdHis') . rand(100, 999);

    //     // Template file ka path
    //     // $templatePath = storage_path('app/template/StudentCertsEnrollment.docx');
    //     $templatePath = storage_path('app/template/Enrollment1.docx');
    //     if (!file_exists($templatePath)) {
    //         dd("File not found at: " . $templatePath);
    //     }

    //   $gender = '';
    //     switch ($student->gender) {
    //         case 1:
    //             $gender = '남성';
    //             break;
    //         case 2:
    //             $gender = '여성';
    //             break;
    //         case 3:
    //             $gender = '기타';
    //             break;
    //         default:
    //             $gender = 'Unknown';
    //     }

    //     // PHPWord TemplateProcessor ko initialize karein
    //     $templateProcessor = new TemplateProcessor($templatePath);

    //     // Placeholders ko replace karein
    //     $templateProcessor->setValue('first_name', $student->last_name);
    //     $templateProcessor->setValue('last_name', $student->first_name);
    //     $templateProcessor->setValue('dob', $student->dob);
    //     $templateProcessor->setValue('nationality', $student->nationality);
    //     $templateProcessor->setValue('gender', $gender);
    //     // $templateProcessor->setValue('date', now()->format('d-m-Y'));
    //     $templateProcessor->setValue('certificate_no', 'CERT-' . str_pad($student->id, 5, '0', STR_PAD_LEFT));
    //     $templateProcessor->setValue('semester', $semesterName->semester_name ?? 'N/A');
    //     $templateProcessor->setValue('semester_start', $semesterName->startDate ?? 'N/A');
    //     $templateProcessor->setValue('semester_end', $semesterName->endDate ?? 'N/A');
    //     // $templateProcessor->setValue('courses_total_hours', $semesterName->subjectHour ?? 'N/A');
    //     $templateProcessor->setValue('program_assigned', $semesterName->programTitle ?? 'N/A');
    //     $templateProcessor->setValue('courses_hours_attendance', $semesterName->totalPresentHours ?? 'N/A');
    //     $templateProcessor->setValue('percentage', $semesterName->attendancePercentage ?? 'N/A');
    //     $templateProcessor->setValue('issuance_number', $issuanceNumber);
        
    //     $templateProcessor->setValue('courses_total_hours', isset($subjectHours[0]) ? $subjectHours[0] : 'N/A');
    //     $templateProcessor->setValue('courses_total_hours_1', isset($subjectHours[1]) ? $subjectHours[1] : 'N/A');
    //     // New file ka naam aur path
    //     $fileName = 'certificate_' . $student->id . '.docx';
    //     $outputPath = storage_path('app/certificates/' . $fileName);

    //     // File ko save karein
    //     $templateProcessor->saveAs($outputPath);

    //     // File ko download karne ka response return karein
    //     return response()->download($outputPath)->deleteFileAfterSend(true);
    // }

public function generateCertificate($id)
    {
   
                  // Student ka data fetch karein
            $student = Student::findOrFail($id);
            
            // Student ki basic information
            $first_name = $student->first_name;
            $last_name = $student->last_name;
            $dob = $student->dob;
            $nationality = $student->nationality; 
            $currentDate = Carbon::now()->format('Y-m-d');
            // Student ki enrollments fetch karein
         $studentEnrolls = StudentEnroll::where('student_id', $student->id)->get();

        // Extract semester IDs from the collection
        $semesterIds = $studentEnrolls->pluck('semester_id');
        
        $studentEnrolls = Semester::whereIn('id', $semesterIds)->get();
        
       
            
         $issuanceNumber = now()->format('YmdHis') . rand(100, 999);

        // Template file ka path
        // $templatePath = storage_path('app/template/StudentCertsEnrollment.docx');
        $templatePath = storage_path('app/template/jhg.docx');
        if (!file_exists($templatePath)) {
            dd("File not found at: " . $templatePath);
        }

      $gender = '';
        switch ($student->gender) {
            case 1:
                $gender = '남성';
                break;
            case 2:
                $gender = '여성';
                break;
            case 3:
                $gender = '기타';
                break;
            default:
                $gender = 'Unknown';
        }

        // PHPWord TemplateProcessor ko initialize karein
        $templateProcessor = new TemplateProcessor($templatePath);




        // Placeholders ko replace karein
        $templateProcessor->setValue('first_name', $first_name);
        $templateProcessor->setValue('last_name', $last_name);
        $templateProcessor->setValue('dob', $dob);
        $templateProcessor->setValue('nationality', $nationality);
        $templateProcessor->setValue('gender', $gender);
       
//       $studentEnrolls = StudentEnroll::where('student_id', $student->id)->get();
// $semesterIds = $studentEnrolls->pluck('semester_id');
// $semesterRecords=Semester::whereIn('id',$semesterIds)->get();

// foreach ($semesterRecords as $semesterRecord) {
//     $templateProcessor->setValue("title", $semesterRecord->title ?? 'N/A');
//     $templateProcessor->setValue("start_date", $semesterRecord->start_date ?? 'N/A');
//     $templateProcessor->setValue("end_date", $semesterRecord->end_date ?? 'N/A');
// }

$studentEnrolls = StudentEnroll::where('student_id', $student->id)->get();
$semesterIds = $studentEnrolls->pluck('semester_id');
$semesterRecords = Semester::whereIn('id', $semesterIds)->get();

// Agar records hain toh table ko dynamically populate karein
if ($semesterRecords->count() > 0) {
    $templateProcessor->cloneRow('index', $semesterRecords->count());
 
    $index = 1;
    foreach ($studentEnrolls as $studentEnroll) {

        // Convert subject_id to an array
        $subjectIds = explode(',', $studentEnroll->subject_id);
        
        // Sum credit hours for this specific enrollment
        $subjectHours = Subject::whereIn('id', $subjectIds)->sum('credit_hour');

        // Get related semester
        $semesterRecord = $semesterRecords->where('id', $studentEnroll->semester_id)->first();

        // Get program title
        $programsTitle = Program::where('semester_id', $studentEnroll->semester_id)->first();

  
            // Get 
        $studentAttendances = StudentAttendance::where('student_id', $studentEnroll->student_id)->first();
        
        // dd($studentAttendances); 
        
        // Set values in Word template
        $templateProcessor->setValue("index#{$index}", $index);
        $templateProcessor->setValue("title#{$index}", $semesterRecord->title ?? 'N/A');
        $templateProcessor->setValue("start_date#{$index}", $semesterRecord->start_date ?? 'N/A');
        $templateProcessor->setValue("end_date#{$index}", $semesterRecord->end_date ?? 'N/A');
        $templateProcessor->setValue("credit_hours#{$index}", $subjectHours ?? 'N/A');
        $templateProcessor->setValue("programs_title#{$index}", $programsTitle->title ?? 'N/A'); // Fix applied here
        $templateProcessor->setValue("present_hours#{$index}", $studentAttendances->present_hours ?? 'N/A');
        $templateProcessor->setValue("total_hours#{$index}", ($studentAttendances->present_hours ?? 0) * 100 / max($subjectHours, 1));

        $templateProcessor->setValue("status#{$index}", $studentEnroll->status); 
 
        $index++;
    }
}



       
        // New file ka naam aur path
        $fileName = 'certificate_' . $student->id . '.docx';
        $outputPath = storage_path('app/certificates/' . $fileName);

        // File ko save karein
        $templateProcessor->saveAs($outputPath);

        // File ko download karne ka response return karein
        return response()->download($outputPath)->deleteFileAfterSend(true);
    }


    
    
    
    public function viewCertificate($id)
    {
        $student = Student::findOrFail($id);
        dd("Partially Working: ", $student);
        $pdfPath = storage_path('app/certificates/certificate_' . $student->id . '.pdf');
    
        if (!file_exists($pdfPath)) {
            return response()->json(['error' => 'Certificate not found'], 404);
        }
    
        return response()->file($pdfPath, [
            'Content-Type' => 'application/pdf',
            'Content-Disposition' => 'inline; filename="certificate.pdf"',
        ]);
    }
    
    
    
    
    
    
    
    
}
