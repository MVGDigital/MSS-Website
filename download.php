<?php
$url='localhost';
$username = "root";
$password = "";
$dbname = "myavtar_livemyavtardb";
$conn = mysqli_connect($url, $username, $password, $dbname);
require_once 'application/libraries/PHPExcel.php';
require_once 'application/libraries/PHPExcel/IOFactory.php';
$result = mysqli_query($conn,"SELECT * FROM master_industry where status='1'");
$result1 = mysqli_query($conn,"SELECT * FROM master_functionality where status='1'");
$cities = mysqli_query($conn,"SELECT * FROM master_cities");
$jobtypes = mysqli_query($conn,"SELECT * FROM master_job_type where status='1'");
$genders = mysqli_query($conn,"SELECT * FROM master_gender where status='1'");
$joblevels = mysqli_query($conn,"SELECT * FROM master_job_level where status='1'");
$noticeperiods = mysqli_query($conn,"SELECT * FROM master_notice_period where status='1'");

$diversities = mysqli_query($conn,"SELECT * FROM master_diversity_group where status='1'");
$orientations = mysqli_query($conn,"SELECT mc.orientation_id,mc.value,mdg.value as diversity FROM master_orientation mc,master_diversity_group mdg where mc.diversity_group_id=mdg.diversity_group_id and mc.status='1'");
$disabilities = mysqli_query($conn,"SELECT mc.disabilty_id,mc.value,mdg.value as diversity FROM master_disabilty mc,master_diversity_group mdg where mc.diversity_group_id=mdg.diversity_group_id and mc.status='1'");
$veterans = mysqli_query($conn,"SELECT mc.veterantype_id,mc.value,mdg.value as diversity FROM master_veterantype mc,master_diversity_group mdg where mc.diversity_group_id=mdg.diversity_group_id and mc.status='1'");
$postretirementtypes = mysqli_query($conn,"SELECT mc.postretirementtype_id,mc.value,mdg.value as diversity FROM master_postretirementtype mc,master_diversity_group mdg where mc.diversity_group_id=mdg.diversity_group_id and mc.status='1'");
$skils = mysqli_query($conn,"SELECT * FROM master_skills where status='1'");
$qualifications = mysqli_query($conn,"SELECT * FROM master_qualification where status='1'");
$courses = mysqli_query($conn,"SELECT mc.course_id,mc.value,mq.value as qualification FROM master_course mc,master_qualification mq where mq.qualification_id = mc.qualification_id and mc.status='1'");
$specializations = mysqli_query($conn,"SELECT mc.specialization_id,mc.value,mq.value as courses FROM master_specialization mc,master_course mq where mq.course_id = mc.course_id and mc.status='1'");
/* Create new PHPExcel object*/
$objPHPExcel = new PHPExcel();

$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(1);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'ID');
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'Value');
$i=2;
while($row1= mysqli_fetch_array($result)) {
	$industry_id=$row1['industry_id'];
	$value=$row1['value'];
	$objPHPExcel->getActiveSheet()->setCellValue("A$i",$industry_id);
	$objPHPExcel->getActiveSheet()->setCellValue("B$i",$value);
$i++;
}
$objPHPExcel->getActiveSheet()->setTitle('Industry');

$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(2);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'ID');
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'Value');
$i=2;
while($row1= mysqli_fetch_array($result1)) {
	$functionality_id=$row1['functionality_id'];
	$value=$row1['value'];
	$objPHPExcel->getActiveSheet()->setCellValue("A$i",$functionality_id);
	$objPHPExcel->getActiveSheet()->setCellValue("B$i",$value);
$i++;
}
$objPHPExcel->getActiveSheet()->setTitle('Funcionality');

$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(3);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'ID');
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'Value');
$i=2;
while($row1= mysqli_fetch_array($cities)) {
	$id=$row1['city_id'];
	$value=$row1['city_name'];
	$objPHPExcel->getActiveSheet()->setCellValue("A$i",$id);
	$objPHPExcel->getActiveSheet()->setCellValue("B$i",$value);
$i++;
}
$objPHPExcel->getActiveSheet()->setTitle('Locations');


$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(4);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'ID');
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'Value');
$i=2;
while($row1= mysqli_fetch_array($jobtypes)) {
	$id=$row1['job_type_id'];
	$value=$row1['value'];
	$objPHPExcel->getActiveSheet()->setCellValue("A$i",$id);
	$objPHPExcel->getActiveSheet()->setCellValue("B$i",$value);
$i++;
}
$objPHPExcel->getActiveSheet()->setTitle('Job Type');


$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(5);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'ID');
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'Value');
$i=2;
while($row1= mysqli_fetch_array($genders)) {
	$id=$row1['gender_id'];
	$value=$row1['value'];
	$objPHPExcel->getActiveSheet()->setCellValue("A$i",$id);
	$objPHPExcel->getActiveSheet()->setCellValue("B$i",$value);
$i++;
}
$objPHPExcel->getActiveSheet()->setTitle('Gender');

$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(6);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'ID');
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'Value');
$i=2;
while($row1= mysqli_fetch_array($joblevels)) {
	$id=$row1['job_level_id'];
	$value=$row1['value'];
	$objPHPExcel->getActiveSheet()->setCellValue("A$i",$id);
	$objPHPExcel->getActiveSheet()->setCellValue("B$i",$value);
$i++;
}
$objPHPExcel->getActiveSheet()->setTitle('Job Level');

$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(7);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'ID');
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'Value');
$i=2;
while($row1= mysqli_fetch_array($noticeperiods)) {
	$id=$row1['notice_period_id'];
	$value=$row1['value'];
	$objPHPExcel->getActiveSheet()->setCellValue("A$i",$id);
	$objPHPExcel->getActiveSheet()->setCellValue("B$i",$value);
$i++;
}
$objPHPExcel->getActiveSheet()->setTitle('Notice Period');

$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(8);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'ID');
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'Value');
$i=2;
while($row1= mysqli_fetch_array($diversities)) {
	$id=$row1['diversity_group_id'];
	$value=$row1['value'];
	$objPHPExcel->getActiveSheet()->setCellValue("A$i",$id);
	$objPHPExcel->getActiveSheet()->setCellValue("B$i",$value);
$i++;
}
$objPHPExcel->getActiveSheet()->setTitle('Diversity Group');

$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(9);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'ID');
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'Diversity');
$objPHPExcel->getActiveSheet()->setCellValue('C1', 'Value');
$i=2;
while($row1= mysqli_fetch_array($orientations)) {
	$id=$row1['orientation_id'];
	$value=$row1['value'];
	$diversity=$row1['diversity'];
	$objPHPExcel->getActiveSheet()->setCellValue("A$i",$id);
	$objPHPExcel->getActiveSheet()->setCellValue("B$i",$diversity);
	$objPHPExcel->getActiveSheet()->setCellValue("C$i",$value);
$i++;
}
$objPHPExcel->getActiveSheet()->setTitle('Orientation');

$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(10);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'ID');
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'Diversity');
$objPHPExcel->getActiveSheet()->setCellValue('C1', 'Value');
$i=2;
while($row1= mysqli_fetch_array($disabilities)) {
	$id=$row1['disabilty_id'];
	$value=$row1['value'];
	$diversity=$row1['diversity'];
	$objPHPExcel->getActiveSheet()->setCellValue("A$i",$id);
	$objPHPExcel->getActiveSheet()->setCellValue("B$i",$diversity);
	$objPHPExcel->getActiveSheet()->setCellValue("C$i",$value);
$i++;
}
$objPHPExcel->getActiveSheet()->setTitle('Disabilty');

$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(11);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'ID');
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'Diversity');
$objPHPExcel->getActiveSheet()->setCellValue('C1', 'Value');
$i=2;
while($row1= mysqli_fetch_array($veterans)) {
	$id=$row1['veterantype_id'];
	$value=$row1['value'];
	$diversity=$row1['diversity'];
	$objPHPExcel->getActiveSheet()->setCellValue("A$i",$id);
	$objPHPExcel->getActiveSheet()->setCellValue("B$i",$diversity);
	$objPHPExcel->getActiveSheet()->setCellValue("C$i",$value);
$i++;
}
$objPHPExcel->getActiveSheet()->setTitle('Veteran Types');

$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(12);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'ID');
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'Diversity');
$objPHPExcel->getActiveSheet()->setCellValue('C1', 'Value');
$i=2;
while($row1= mysqli_fetch_array($postretirementtypes)) {
	$id=$row1['postretirementtype_id'];
	$value=$row1['value'];
	$diversity=$row1['diversity'];
	$objPHPExcel->getActiveSheet()->setCellValue("A$i",$id);
	$objPHPExcel->getActiveSheet()->setCellValue("B$i",$diversity);
	$objPHPExcel->getActiveSheet()->setCellValue("C$i",$value);
$i++;
}
$objPHPExcel->getActiveSheet()->setTitle('Post Retirement Types');

$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(13);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'ID');
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'Value');
$i=2;
while($row1= mysqli_fetch_array($skils)) {
	$id=$row1['skills_id'];
	$value=$row1['value'];
	$objPHPExcel->getActiveSheet()->setCellValue("A$i",$id);
	$objPHPExcel->getActiveSheet()->setCellValue("B$i",$value);
$i++;
}
$objPHPExcel->getActiveSheet()->setTitle('Skills');

$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(14);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'ID');
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'Value');
$i=2;
while($row1= mysqli_fetch_array($qualifications)) {
	$id=$row1['qualification_id'];
	$value=$row1['value'];
	$objPHPExcel->getActiveSheet()->setCellValue("A$i",$id);
	$objPHPExcel->getActiveSheet()->setCellValue("B$i",$value);
$i++;
}
$objPHPExcel->getActiveSheet()->setTitle('Qualifications');

$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(15);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'ID');
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'Qualification');
$objPHPExcel->getActiveSheet()->setCellValue('C1', 'Value');
$i=2;
while($row1= mysqli_fetch_array($courses)) {
	$id=$row1['course_id'];
	$value=$row1['value'];
	$qualification=$row1['qualification'];
	$objPHPExcel->getActiveSheet()->setCellValue("A$i",$id);
	$objPHPExcel->getActiveSheet()->setCellValue("B$i",$qualification);
	$objPHPExcel->getActiveSheet()->setCellValue("C$i",$value);
$i++;
}
$objPHPExcel->getActiveSheet()->setTitle('Courses');

$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(16);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'ID');
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'Qualification');
$objPHPExcel->getActiveSheet()->setCellValue('C1', 'Value');
$i=2;
while($row1= mysqli_fetch_array($specializations)) {
	$id=$row1['specialization_id'];
	$value=$row1['value'];
	$courses=$row1['courses'];
	$objPHPExcel->getActiveSheet()->setCellValue("A$i",$id);
	$objPHPExcel->getActiveSheet()->setCellValue("B$i",$courses);
	$objPHPExcel->getActiveSheet()->setCellValue("C$i",$value);
$i++;
}
$objPHPExcel->getActiveSheet()->setTitle('Specializations');


$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->setCellValue('A1', 'Job Title');
$objPHPExcel->getActiveSheet()->setCellValue('B1', 'Industry');
$objPHPExcel->getActiveSheet()->setCellValue('C1', 'Functional Area');
$objPHPExcel->getActiveSheet()->setCellValue('D1', 'Location Cities');
$objPHPExcel->getActiveSheet()->setCellValue('E1', 'Job Type');
$objPHPExcel->getActiveSheet()->setCellValue('F1', 'No Of Vacancies');
$objPHPExcel->getActiveSheet()->setCellValue('G1', 'Gender');
$objPHPExcel->getActiveSheet()->setCellValue('H1', 'Level');
$objPHPExcel->getActiveSheet()->setCellValue('I1', 'Experience From');
$objPHPExcel->getActiveSheet()->setCellValue('J1', 'Experience To');
$objPHPExcel->getActiveSheet()->setCellValue('K1', 'Notice Period');
$objPHPExcel->getActiveSheet()->setCellValue('L1', 'Diversity');
$objPHPExcel->getActiveSheet()->setCellValue('M1', 'Orientation');
$objPHPExcel->getActiveSheet()->setCellValue('N1', 'Disability');
$objPHPExcel->getActiveSheet()->setCellValue('O1', 'Veteran Type');
$objPHPExcel->getActiveSheet()->setCellValue('P1', 'Post Retirement Type');
$objPHPExcel->getActiveSheet()->setCellValue('Q1', 'Job Description');
$objPHPExcel->getActiveSheet()->setCellValue('R1', 'Skills');
$objPHPExcel->getActiveSheet()->setCellValue('S1', 'Qualtification');
$objPHPExcel->getActiveSheet()->setCellValue('T1', 'Course');
$objPHPExcel->getActiveSheet()->setCellValue('U1', 'Specialization');
$objPHPExcel->getActiveSheet()->setCellValue('V1', 'Additional Requirement');
$objPHPExcel->getActiveSheet()->setCellValue('W1', 'Annual CTC From');
$objPHPExcel->getActiveSheet()->setCellValue('X1', 'Annual CTC To');
$objPHPExcel->getActiveSheet()->setCellValue('Y1', 'Hide Salary');
$objPHPExcel->getActiveSheet()->setCellValue('Z1', 'First Posted On');
$objPHPExcel->getActiveSheet()->setCellValue('AA1', 'Last Date to Apply');
$objPHPExcel->getActiveSheet()->setCellValue('AB1', 'Company Profile');
//First row sample value
$objPHPExcel->getActiveSheet()->setCellValue('A2', 'Specify Your Job Title');
$objPHPExcel->getActiveSheet()->setCellValue('B2', 'Enter ID from sheet "Industry"');
$objPHPExcel->getActiveSheet()->setCellValue('C2', 'Enter ID from sheet "Funcionality"');
$objPHPExcel->getActiveSheet()->setCellValue('D2', 'Enter ID from sheet "Locations"');
$objPHPExcel->getActiveSheet()->setCellValue('E2', 'Enter ID from sheet "Job Type"');
$objPHPExcel->getActiveSheet()->setCellValue('F2', 'Enter No Of Vacancies (Number only)');
$objPHPExcel->getActiveSheet()->setCellValue('G2', 'Enter ID from sheet "Gender"');
$objPHPExcel->getActiveSheet()->setCellValue('H2', 'Enter ID from sheet "Job Level"');
$objPHPExcel->getActiveSheet()->setCellValue('I2', 'Experience From');
$objPHPExcel->getActiveSheet()->setCellValue('J2', 'Experience To');
$objPHPExcel->getActiveSheet()->setCellValue('K2', 'Enter ID from sheet "Notice Period"');
$objPHPExcel->getActiveSheet()->setCellValue('L2', 'Enter ID from sheet "Diversity Group"');
$objPHPExcel->getActiveSheet()->setCellValue('M2', 'Enter ID from sheet "Orientation"');
$objPHPExcel->getActiveSheet()->setCellValue('N2', 'Enter ID from sheet "Disabilty"');
$objPHPExcel->getActiveSheet()->setCellValue('O2', 'Enter ID from sheet "Veteran Types"');
$objPHPExcel->getActiveSheet()->setCellValue('P2', 'Enter ID from sheet "Post Retirement Types"');
$objPHPExcel->getActiveSheet()->setCellValue('Q2', 'Enter Job Description');
$objPHPExcel->getActiveSheet()->setCellValue('R2', 'Enter ID from sheet "Skills"');
$objPHPExcel->getActiveSheet()->setCellValue('S2', 'Enter ID from sheet "Qualtifications"');
$objPHPExcel->getActiveSheet()->setCellValue('T2', 'Enter ID from sheet "Course"');
$objPHPExcel->getActiveSheet()->setCellValue('U2', 'Enter ID from sheet "Specialization"');
$objPHPExcel->getActiveSheet()->setCellValue('V2', 'Additional Requirement');
$objPHPExcel->getActiveSheet()->setCellValue('W2', 'Annual CTC From');
$objPHPExcel->getActiveSheet()->setCellValue('X2', 'Annual CTC To');
$objPHPExcel->getActiveSheet()->setCellValue('Y2', 'Hide Salary');
$objPHPExcel->getActiveSheet()->setCellValue('Z2', 'Date Format "2022-11-15"');
$objPHPExcel->getActiveSheet()->setCellValue('AA2', 'Date Format "2022-11-15"');
$objPHPExcel->getActiveSheet()->setCellValue('AB2', 'Company Profile');

$objPHPExcel->getActiveSheet()->setTitle('Job Details');

header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="job_uplaod_template-'. date("Y-m-d").'.xls"');
header('Cache-Control: max-age=0');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');