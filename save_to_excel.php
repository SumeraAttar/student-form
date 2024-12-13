<?php
if ($_SERVER["REQUEST_METHOD"] == "POST") {
    // Retrieve form data
    $fullName = $_POST['fullName'];
    $username = $_POST['username'];
    $email = $_POST['email'];
    $phoneNumber = $_POST['phoneNumber'];
    $domin = $_POST['Domin'];
    $expFr = $_POST['Exp/fr'];
    $companyName = $_POST['cn'];
    $yearsOfExperience = $_POST['ey'];
    $designation = $_POST['desi'];
    $ctc = $_POST['ctc'];
    $gender = isset($_POST['gender']) ? $_POST['gender'] : 'Not Specified';

    // File path to store the Excel file
    $filePath = "submissions.xlsx";

    // Include PHPExcel library
    require 'PHPExcel.php';

    // Load existing file or create a new one
    if (file_exists($filePath)) {
        $spreadsheet = PHPExcel_IOFactory::load($filePath);
    } else {
        $spreadsheet = new PHPExcel();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A1', 'Full Name')
              ->setCellValue('B1', 'City')
              ->setCellValue('C1', 'Email')
              ->setCellValue('D1', 'Phone Number')
              ->setCellValue('E1', 'Domain')
              ->setCellValue('F1', 'Experience/Fresher')
              ->setCellValue('G1', 'Company Name')
              ->setCellValue('H1', 'Years of Experience')
              ->setCellValue('I1', 'Designation')
              ->setCellValue('J1', 'CTC')
              ->setCellValue('K1', 'Gender');
    }

    // Get the active sheet
    $sheet = $spreadsheet->getActiveSheet();

    // Find the next empty row
    $row = $sheet->getHighestRow() + 1;

    // Write data to the next row
    $sheet->setCellValue("A$row", $fullName)
          ->setCellValue("B$row", $username)
          ->setCellValue("C$row", $email)
          ->setCellValue("D$row", $phoneNumber)
          ->setCellValue("E$row", $domin)
          ->setCellValue("F$row", $expFr)
          ->setCellValue("G$row", $companyName)
          ->setCellValue("H$row", $yearsOfExperience)
          ->setCellValue("I$row", $designation)
          ->setCellValue("J$row", $ctc)
          ->setCellValue("K$row", $gender);

    // Save the file
    $writer = PHPExcel_IOFactory::createWriter($spreadsheet, 'Excel2007');
    $writer->save($filePath);

    echo "Data saved successfully!";
}
?>
