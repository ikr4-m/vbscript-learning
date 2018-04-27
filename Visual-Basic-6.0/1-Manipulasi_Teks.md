# Manipulasi Teks

Ok, kita akan mencoba dasarnya dulu. Kita akan coba memodifikasi sebuah Label dengan menggunakan beberapa tombol.

![Manipulasi Teks](https://media.discordapp.net/attachments/308790604723257365/439067544809111552/VB6_2018-04-26_22-17-01.png)

Kira-kira seperti inilah form yang akan kita buat, untuk nama tombolnya, kalian bisa modifikasi seperti __Command1, kalian bisa ganti dengan tblInput dengan cara seperti di bawah:__

![tblInput](https://media.discordapp.net/attachments/308790604723257365/439068380205678592/VB6_2018-04-26_22-20-53.png)

Disini, kalian akan mengenal fungsi untuk:
* Membuat teks menjadi:
  * Tebal
  * Miring
  * Bergaris bawah
* Menonaktifkan atau menyalakan tombol atau CommandButton
* Mengubah warna teks
* Fungsi pengulangan sederhana (If..Then..Else)

Nah, berikut adalah listingnya. Nanti saya berikan penjelasan di tiap-tiap barisnya:

```vb
Private Sub Check1_Click()          ' Kondisi ketika Check1 diklik
Label2.FontBold = Check1.Value      ' Apabila Check1 dipilih, maka Label2 akan tercetak tebal
End Sub

Private Sub Check2_Click()          ' Kondisi ketika Check2 diklik
Label2.FontItalic = Check2.Value    ' Apabila Check2 diklik, maka Label2 akan tercetak miring
End Sub

Private Sub Check3_Click()            ' Kondisi ketika Check3 diklik
Label2.FontUnderline = Check3.Value   ' Apabila Check3 diklik, maka Label2 akan tercetak dengan garis bawah
End Sub

Private Sub Command1_Click()          ' Kondisi ketika menekan Command1
Label2.Caption = Text1.Text           ' Label2 akan menjiplak isi dari Text1
Command3.Enabled = True               ' Tombol Command3 dinyalakan
Text1.Text = ""                       ' Tanda "" menandakan bahwa Text1 dikosongkan
End Sub

Private Sub Command2_Click()          ' Kondisi ketika menekan Command2
Unload Me                             ' Sama seperti tombol Close pada umumnya, biasanya tertulis End atau Unload Me
End Sub

Private Sub Command3_Click()          ' Kondisi ketika menekan Command3
Label2.Caption = ""                   ' Mengosongkan isi Label2
Label2.ForeColor = vbBlack            ' Mengubah warna Label2 menjadi warna vbBlack atau Hitam
End Sub

Private Sub Form_Load()               ' Kondisi ketika Form baru dijalankan
Text1.Text = ""                       ' Mengosongkan isi Text1
Label2.Caption = ""                   ' Mengosongkan isi Label2
Command1.Enabled = False              ' Tombol Command1 dinonaktifkan
Command3.Enabled = False              ' Tombol Command2 dinonaktifkan
End Sub

Private Sub Option1_Click()           ' Kondisi ketika Option1 dipilih
Label2.ForeColor = vbBlue             ' Mengubah warna Label2 menjadi warna vbBlue atau Biru
End Sub

Private Sub Option2_Click()           ' Kondisi ketika Option2 dipilih
Label2.ForeColor = vbRed              ' Mengubah warna Label2 menjadi warna vbRed atau Merah
End Sub

Private Sub Option3_Click()           ' Kondisi ketika Option3 dipilih
Label2.ForeColor = vbGreen            ' Mengubah warna Label2 menjadi warna vbGreen atau Hijau
End Sub

Private Sub Text1_Change()            ' Kondisi ketika Text1 berubah walau satu karakter pun

' Ok kali ini kita akan bermain logika, jadi disini contohnya seperti ini.
' Apabila saya menuliskan walau satu karakter di Text1, maka tombol Command1 dinyalakan atau dapat ditekan.
' Nah dari logika itu kita dapat buatkan listingnya seperti berikut

If Text1.Text = "" Then               ' Meminta kondisi Text1 ketika kosong
    Command1.Enabled = False          ' Apabila kondisi diatas benar, maka Command1 dinonaktifkan
    Else                              ' Kata Else mewakili selain dari kondisi diatasnya
        Command1.Enabled = True       ' Command1 dinyalakan
    End If                            ' Jangan lupa untuk menutup fungsi pengulangannya
End Sub
```
