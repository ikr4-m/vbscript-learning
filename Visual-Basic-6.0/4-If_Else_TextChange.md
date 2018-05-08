# Database Sederhana

*Jujur, ini adalah bagian yang paling gua suka*

Ok, nah kali ini kita akan coba-coba membuat database sederhana!

![Database Sederhana](https://cdn.discordapp.com/attachments/439384229483380746/443403501473628180/unknown.png)

Logikanya gini, jadi nanti disini ada **satu output dan tiga output**. __Output__ itu adalah **field Nama Siswa, Tahun, dengan Jurusan** sedangkan __Input__ itu adalah **field NIS doang**.

![Contoh](https://cdn.discordapp.com/attachments/439384229483380746/443415102243471361/unknown.png)

### Studi Kasus

Database itu kan __perkumpulan data yang disatukan untuk memudahkan seseorang untuk memanajemen sebuah data__ *menurut gua*. Berarti, kita perlu beberapa sampel data untuk dijadikan *kelinci percobaan*.

NIS | Nama Siswa | Tahun | Jurusan
----|------------|-------|--------
172-091 | Andi Muh. Syahrul | 2017 | RPL
172-002 | Adrian Pratama | 2017 | RPL
161-001 | Andi Muh. Akbar | 2016 | TKJ
161-002 | Abdullah Faqih Mustofa | 2016 | TKJ

Ok, dari tabel diatas, kita mengetahui:

1. Apabila dua angka di depan NIS memenuhi persyaratan berikut:
  1. "16" maka field Tahun adalah 2016
  2. "17" maka field Tahun adalah 2017
```vb
Left(String, Length As Long)
```
2. Apabila karakter ketiga dari kiri (sebelah kiri garis datar) memenuhi persayaratan berikut:
  1. "1" maka field Jurusan adalah TKJ
  2. "2" maka field Jurusan adalah RPL
```vb
Mid(String, Start As Long, [Length])
```
3. Dan untuk field namanya, *beda orang beda nomor urut toh...*

## Tujuan Pembelajaran

Di sini kalian akan belajar fungsi tentang:
* Pengambilan String (Santai, ini macam fungsi =LEFT, =MID, dengan =RIGHT di Excel kok)
  * Dari sebelah kiri
  * Dari tengah
* Pengulangan **IF**
  * Penggunaan ElseIf
  * Mengenalkan Value/String dalam sebuah teks

### Listing

```vb
Private Sub Form_Load()
'Mengosongkan semua textbox
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""

'Biar app tidak galat/rusak, textbox yang terisi otomatis dimatikan
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
End Sub

Private Sub Text1_Change()
' Kalian mungkin bertanya-tanya mengapa kita cuma fokus di Text1 doang?
' 1. Karena dalang dari perubahan semua field itu berasal dari Text1
' 2. Logikanya, anda ga bakal gerak kalau ibu anda belum ngasih perintah.

'Untuk ngubah Text3 secara otomatis
If Left(Text1, 2) = "17" Then
    Text3 = "2017"
Else
    Text3 = ""
    End If

'Untuk ngubah Text4 secara otomatis
'ElseIf digunakan untuk mengurangi penutupan End If..End If yang terlalu banyak
If Mid(Text1, 3, 1) = "1" Then
    Text4 = "Teknik Komputer Jaringan"
ElseIf Mid(Text1, 3, 1) = "2" Then
    Text4 = "Rekayasa Perangkat Lunak"
Else
    Text4 = ""
    End If

'Untuk ngubah Text2 secara otomatis
'ElseIf digunakan untuk mengurangi penutupan End If..End If yang terlalu banyak
'Pakai "xxx-xxx" biar nanti kalau belakangnya kembar bakal gabisa tertabrak
'Kemudian, gapake Right,Left,Right function karena kepanjangan
If Text1 = "172-001" Then
    Text2 = "Andi Muh. Syahrul"
ElseIf Text1 = "172-002" Then
    Text2 = "Adrian Pratama"
ElseIf Text1 = "161-001" Then
    Text2 = "Andi Muh. Akbar"
ElseIf Text1 = "161-002" Then
    Text2 = "Abdullah Faqih Mustofa"
Else
    Text2 = ""
    End If
End Sub
```
