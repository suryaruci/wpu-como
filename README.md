# wpu-como
Repository untuk menyimpan scripting COMO
Create New file Excel with sheet Verbatim (Verbatim want to Code) and Sheet Info (Information of questions verbatim)
1. Sheet Info
  1.1. "Click me" Button
    1.1.1. Import verbatim (On Progress)
      1.1.1.1. Deskription : Membuat sheet verbatim yang isinya adalah hasil copy isi sheet verbatim di file Panter
      1.1.1.2. Cara Kerja : 
               1. Pilih tombol Click me… (pojok kiri atas)
               2. Pilih import verbatim
               3. Akan muncul dialog box browse file, pilih file panter yang akan di ambil verbatimnya
               4. Akan muncul sheet verbatim hasil copy verbatim yang ada di difile panter
    1.1.2. Questions Info (Only if Sheet Verbatim exist)
      1.1.2.1. Deskription : Menampilkan informasi dari pertanyaan-pertanyaan yang akan di Coding (Ada di table Project)
      1.1.2.2. Cara Kerja :
               1. Project -->Isi nama projectnya
               2. No--> Nomor urut/indexing
               3. QuestID-->Pertanyaan yang akan di coding
               4. Quest Label --> Label pertanyaan
               5. N-Verb-->Jumlah verbatim dalam suatu pertanyaan
               6. Weight-->Bobot pertanyaan secara keseluruhan
               7. Complexity-->Tingkat kesulitan pertanyaan
                  a. Very easy-->Verbatim pendek(brand)/kurang lebih 1-2 code
                  b. Easy-->Verbatim pendek(others)/kurang lebih 3-5 code
                  c. Middle -->Verbatimnya sedang/kurang lebih 6-9 code
                  d. Hight-->Jawaban lebih 10 code 
               8. Assigned-->Nama coder yang diassign untuk pertanyaan tersebut
               9. Frame-->Membuat template Frame yang akan digunakan (Klik kanan di garis merah Kolom Frame):
                  a. Go to Question-->ke pertanyaan di sheet Data berdasarkan question di column QuetionID di Sheet Info 
                  b. Create Frame-->Membuat template Frame sesuai nama frame di coulum yang aktif/dipilih
                  c. Go to Frame-->Membuka Frame sesuai nama frame di coulmn yang aktif/dipilih
                  d. Refresh Coder Name-->Menambahkan/Merubah nama coder
               10. Coded (%)-->Menampilkan pengkodingan oleh coder (%)"
    1.1.3. Transpose verbatim
      1.1.3.1. Deskription : Menyusun/transpose verbatim supaya memudahkan coder dalam melakukan coding
      1.1.3.2. Cara Kerja :
               1. All Verbatim-->Transpose semua pertanyaan
               2. (Nama Coder)-->Transpose hanya pertanyaan yang di Assign (name of Coder must be assign in Column Assigned)
    1.1.4. Productivity
      1.1.4.1. Deskription : Menampilkan persentase pekerjaan/coding dari yang dilakukan oleh coder (Table Productivity)
      1.1.4.1. Cara Kerja : Choose productivity
    1.1.5. Back to Field
      1.1.5.1. Deskription : Verbatim yang perlu di kembalikan ke Field untuk di tanyakan kembali maksud atau lebih detail ke responden
      1.1.5.2. Cara Kerja : Berikan warna merah pada Kalimat yang perlu di kembalikan ke field yang kurang Probing/tidak relevan dll, kemudian di Column Note di tulis keterangannya yaitu kurang Probing/tidak relevan dll, sesuai dengan maksud dikembalikan
    1.1.6. Verification summary (on Progress)
2. Sheet Verbatim : Hasil copy sheet verbatim dari file panter
3. Sheet Data
  3.1. Header
      1. Serial --> Nomor serial/responden
      2. Quest-->Pertanyaan
      3. Verbatim-->Isi dari pertanyaan
      4. Coding-->Column untuk mengcoding
      5. Search-->Colum untuk menulis keyword untuk mencari pilihan jawaban dari frame yang di maksud
      6. Code-->Menampilkan Code dari codelist yang muncul saat menemukan keyword di frame
      7. Codelist-->Menampilkan pilihan jawaban di frame sesuai keyword yang di masukan di column search
      8. Transfer Code-->Hasil transfer coder ID ke Client ID (Jika pilih transfer dan Export to CSV File di Tombol Menu)
      9. Verification-->Menampilkan gabungan Code dan Codelist dari Frame sesuai pertanyaan yang akan di gunakan untuk membandingkan dengan hasil Coding
      10. Coder-->Nama Coder yang mengcoding pertanyaan tersebut
      11. Verificator-->Orang yang akan memverifikasi
      12. Information-->Info if error
      13. ID INTV-->ID Interviewer untuk pertanyaan tersebut
      14. City-->Kota responden yang menjawab pertanyaan tersebut
      15. Note-->Trigger for verbatim (for Back to field/Query dll)
      16. Index-->Nomor urut
      17. Identity-->Gabungan antara Serial dan Question"
  3.2."To Frame" Button
    3.2.1. Deskription : Jump to frame base on Question in Column Quest (Data)
    3.2.2. Cara Kerja : Arahkan cursor ke column D, maka frame yang di gunakan oleh pertanyaan di Column B akan Active
  3.3. "Menu" Button
    3.3.1. Exact Coding
      3.3.1.1. Deskription : Automatic coding untuk brand
      3.3.1.2. Cara kerja :
              1. Pilih Exact Coding
              2. Pilih frame yang akan di gunakan acuan dalam coding
              3. Hasil coding otomatis akan keluar di Column Coding, jika kosong makan jawaban/verbatim tidak ada di frame atau ada typo (Verbatim dan Frame harus benar-benar sama/Exact)
     3.3.2. Similarity Coding (On Progress)
       3.3.2.1. Deskription : Automatic menampilkan semua kemungkinan hasil coding untuk verbatim panjang, coder memilih yang tepat/merubahnya dengan yang sesuai
       3.3.2.2. Cara Kerja :
              1. Pilih Similarity Coding
              2. Pilih frame yang akan di gunakan acuan dalam coding-->Kemungkinan hasil coding pilihan program akan keluar di Column Coding
    3.3.3. Search
      3.3.3.1. Deskription : Menampilan pilihan jawaban yang sesuai dengan keyword yang di masukan
      3.3.3.2. Cara Kerja :
              1. Pilih Seach
              2. Pilih frame yang akan digunakan acuan dalam coding
              3. Masukan keyword yang di cari di Column Search, akan muncul pilihan jawaban di column Codelist dan CodeID di column Code
              4. Pilih Code di column Code yang sesuai dengan verbatim  lalu Tekan Ctrl+M untuk memindahkan code ke column Coding
              5. Pilih Cell/Range di Column Coding dengan verbatim yang makananya sama/semisal lalu Tekan Ctrl+L untk mengulang coding yang sudah dilakukan dipoint 4
              6. Pilih Cell/Range di Column Coding yang akan di delete/undo code yang terakhir di coding lalu tekan Ctrl+Shift+L
    3.3.4. Run Report
      3.3.4.1. Deskription : Meampilkan error/codingan yang blank, diluar frame
      3.3.4.2. Cara Kerja :
              1. Pilih Run Report
              2. Pilih frame yang akan dicek, Total error in Message Box, Information error in Column information
    3.3.5. Verification
      3.3.5.1. Deskription : Memunculkan gabungan kode dan Frame di kolomn Verifikasi sesuai dengan frame
      3.3.5.2. Cara Kerja :
              1. Pilih Verification
              2. Pilih frame yang akan diverifikasi, Muncul di column verifikasi gabungan Code dan Codelist dan secara otomatis akan muncul 10% verbatim yang harus di verifikasi Coder (Warna biru)
    3.3.6. Create Identity
      3.3.6.1. Deskription : Gabungan antara Serial dan Question, digunakan untuk penggabungan file coding
      3.3.6.2. Cara Kerja : Pilih Create Identity-->akan muncul di column verifikasi gabungan Serial dan Question di Column Identity
    3.3.7. Export to CSV File
      3.3.7.1. Deskription : Create sheet UpdateCSV (CSV data format) after coding verbatim finished
      3.3.7.2. Cara Kerja :
              1. Pilih Export to CSV File-->Membuat sheet Update CSV yang siap di kirim ke Programmer 
              2. File "UpdateCSV-NameFile" will be created automaticly and save in Same folder Project
    3.3.8. Transfer and Export to CSV File
      3.3.8.1. Deskription : Create sheet UpdateCSV (CSV data format) after coding verbatim finished yang sudah ditransfer
      3.3.8.2. Cara Kerja : 
              1. Pilih Transfer and Export to CSV File
              2. Pilih Frame yang akan di transfer/All Frame will be created automaticly and save in Same folder Project 
4. Sheet Frame
  4.1. Header
       1. Quest-->Pertanyaan-pertanyaan yang menggunakan frame tersebut
       2. CoderID-->Code awal/code yang di ajukan coder
       3. ClientID-->Code yang di Approve oleh Client/Researcher
       4. Statement (Bahasa)-->Codelist/frame dalam bahasa
       5. Statement (English)-->Coelist/frame dalam English
       6. Note-->Untuk catatan/yang perlu di perhatikan
       7. Informastion-->Column untuk cek error
       8. Count-->Frekuensi dari suatu frame/codelist (Secara Total), di samping akan muntul detail perpertanyaan"
       9. Count aech Question (if Quest more than one)
  4.2. "Back" Button : Membuka/mengaktifkan sheet Data
  4.3. "Menu" Button
    4.3.1. Frequesncy
      4.3.1.1. Deskription : Mengeluarkan count masing-masing pertanyaan
      4.3.1.2. Cara Kerja : Pilih Frequency-->akan keluar Count Total dan per pertanyaan di column paling kanan 
    4.3.2. Run Report
      4.3.2.1. Deskription : Meampilkan error/frame baru tetapi tidak digunakan di codingan
      4.3.2.2. Cara Kerja : Pilih Run report-->akan keluar keterangan error di column Information jika ada
    4.3.3. Create Query (On Progress)
5. Sheet UpdateCSV : File data hasil coding yang digunakan oleh programmer/Client
6. Sheet BactToField : File for Back to field
  6.1. Header
      1. Project : Name project
      2. No--> Number
      3. ID INTV --> ID Interviewer
      4. Serial --> Serial question
      5. Quest --> Question
      6. Verbatim --> Original verbatim contains statement which back to Field to probe/re-do etc
      7. Concern --> statement which back to Field to probe/re-do etc
      8. Note-->Note for Concern statemen
      9. Confirm From Field-->Feedback from field to revise or clarification
      10. Code --> Coding concern statemen
  6.2. "Adding code" button : Value/code in colum code auto transfer to Column coding in sheet Data
