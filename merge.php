<?php

use PDFMerger\PDFMerger;
require_once 'PDFMerger.php';

$pdf = new PDFMerger();


$pdf->addPDF('D:\xampp\htdocs\projects\myavtar\asset\resume\topsheet\topsheet-A Gomathi1623229418.pdf', 'all');
$pdf->addPDF('D:\xampp\htdocs\projects\myavtar\asset\resume\00a93f0241718030bdcba0539b11f865.pdf', 'all');


$pdf->merge('file', 'D:\xampp\htdocs\projects\myavtar\asset\resume\mew.pdf'); // generate the file

$pdf->merge('download', 'D:\xampp\htdocs\projects\myavtar\asset\resume\mew.pdf'); // force download


?>