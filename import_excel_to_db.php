<?php
// Load file koneksi.php
include "koneksi.php";

// Load file autoload.php
require 'vendor/autoload.php';

// Include librari PhpSpreadsheet
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

if(isset($_POST['import'])){ // Jika user mengklik tombol Import
  $nama_file_baru = $_POST['namafile'];
    $path = 'tmp/' . $nama_file_baru; // Set tempat menyimpan file tersebut dimana

    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    $spreadsheet = $reader->load($path); // Load file yang tadi diupload ke folder tmp
    $sheet = $spreadsheet->getActiveSheet()->toArray(null, true, true, true);

  $numrow = 1;
  foreach($sheet as $row){
    // Ambil data pada excel sesuai Kolom
    $nis = $row['A']; // Ambil data NIS
    $nama = $row['B']; // Ambil data nama
    $jenis_kelamin = $row['C']; // Ambil data jenis kelamin
    $telp = $row['D']; // Ambil data telepon
    $alamat = $row['E']; // Ambil data alamat

    // Cek jika semua data tidak diisi
    if($nis == "" && $nama == "" && $jenis_kelamin == "" && $telp == "" && $alamat == "")
      continue; // Lewat data pada baris ini (masuk ke looping selanjutnya / baris selanjutnya)

    // Cek $numrow apakah lebih dari 1
    // Artinya karena baris pertama adalah nama-nama kolom
    // Jadi dilewat saja, tidak usah diimport
    if($numrow > 1){
      // Buat query Insert
      $query = "INSERT INTO siswa VALUES('" . $nis . "','" . $nama . "','" . $jenis_kelamin . "','" . $telp . "','" . $alamat . "')";

      // Eksekusi $query
      mysqli_query($connect, $query);
    }

    $numrow++; // Tambah 1 setiap kali looping
  }

    unlink($path); // Hapus file excel yg telah diupload, ini agar tidak terjadi penumpukan file
}

header('location: index.php'); // Redirect ke halaman awal