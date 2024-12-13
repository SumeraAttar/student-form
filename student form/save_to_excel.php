<?php
require 'vendor/autoload.php'; // Include PhpSpreadsheet library if using Composer
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// File to store data
$filePath = 'form-data.xlsx';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    // Collect form data
    $formData = [
        'Full Name' => $_POST['fullName'] ?? '',
        'City' => $_POST['username'] ?? '',
        'Email' => $_POST['email'] ?? '',
        'WhatsApp Number' => $_POST['phoneNumber'] ?? '',
        'Domain' => $_POST['Domin'] ?? '',
        'Experience/Fresher' => $_POST['Exp/fr'] ?? '',
        'Company Name' => $_POST['cn'] ?? '',
        'Years of Experience' => $_POST['ey'] ?? '',
        'Designation' => $_POST['desi'] ?? '',
        'CTC' => $_POST['ctc'] ?? '',
        'Gender' => $_POST['gender'] ?? ''
    ];

    // Load or create an Excel file
    if (file_exists($filePath)) {
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($filePath);
        $sheet = $spreadsheet->getActiveSheet();
    } else {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->fromArray(array_keys($formData), NULL, 'A1'); // Add headers
    }

    // Append form data to the sheet
    $sheet->fromArray(array_values($formData), NULL, 'A' . ($sheet->getHighestRow() + 1));

    // Save the Excel file
    $writer = new Xlsx($spreadsheet);
    $writer->save($filePath);

    echo 'Form data saved successfully!';
} else {
    echo 'Invalid request method.';
}
?>
