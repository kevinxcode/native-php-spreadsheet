<?php
// Load file koneksi.php
include "conn.php";

// Load file autoload.php
require 'vendor/autoload.php';

// Include librari PhpSpreadsheet
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;


if(isset($_POST['import'])){ // Jika user mengklik tombol Import
  // ambil data file
    $namaFile = $_FILES['namafile']['name'];
    $namaSementara = $_FILES['namafile']['tmp_name'];
    $ext = pathinfo($namaFile, PATHINFO_EXTENSION);
  	$namabaru = "excel_import.".$ext;
    // tentukan lokasi file akan dipindahkan
    $dirUpload = "tmp/";
    $terupload = move_uploaded_file($namaSementara, $dirUpload.$namabaru);

    // $nama_file_baru = 'a.xlsx';
    $path = 'tmp/' . $namabaru; // Set tempat menyimpan file tersebut dimana

    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    $spreadsheet = $reader->load($path); // Load file yang tadi diupload ke folder tmp
    $sheet = $spreadsheet->getActiveSheet()->toArray(null, true, true, true);

  $numrow = 1;
  // print_r($sheet); exit;
  foreach($sheet as $row){
    // Ambil data pada excel sesuai Kolom
    $id = $row['A']; // Ambil data nama
    $nama = $row['B']; // Ambil data nama
    $kelas = $row['C']; // Ambil data kelas
    $alamat = $row['D']; // Ambil data jenis alamat
    // $query = "INSERT INTO tb_siswa VALUES('" . $id . "','" . $nama . "','" . $kelas . "','" . $alamat . "')";
    // mysqli_query($koneksi, $query);
    // Cek jika semua data tidak diisi
    if($nama == "" && $kelas == "" && $alamat == "")
      continue; // Lewat data pada baris ini 
    // Cek $numrow apakah lebih dari 1
    // Artinya karena baris pertama adalah nama-nama kolom
    // Jadi dilewat saja, tidak usah diimport
    if($numrow > 1){
      // Buat query Insert
      $query = "INSERT INTO tb_siswa VALUES('" . $id . "','" . $nama . "','" . $kelas . "','" . $alamat . "')";
      mysqli_query($koneksi, $query);
    }
    $numrow++; // Tambah 1 setiap kali looping
  }

    unlink($path); // Hapus file excel yg telah diupload, ini agar tidak terjadi penumpukan file
}

header('location: index.php'); // Redirect ke halaman awal