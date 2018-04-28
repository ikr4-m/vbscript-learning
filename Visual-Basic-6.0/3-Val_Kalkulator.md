# Kalkulator Sederhana

*Walau kalkulator ini ga macam kalkulator di kompi, tapi yang penting bisa ngitung cepat kan ehehe*
Nah, kali ini kita akan membuat kalkulator sederhana!

![Kalkulator](https://media.discordapp.net/attachments/439384229483380746/439625208794578944/VB6_2018-04-28_11-13-18.png)

*Ini kalkulator cuma bisa ngitung 2 variabel doang btw, gpp yang penting bisa ngitung.* Oh iya, ini kalkulator dapat menggunakan operasi logika seperti yang tertera pada frame di sebelah kiri itu. Nanti hasil yang muncul dari operasi logika bukan angka tetapi ~~hukum CLI~~ melainkan huruf `True` ataupun `False`.

![Tru en Fals](https://media.discordapp.net/attachments/439384229483380746/439637329158864897/VB6_2018-04-28_12-01-40.png)

## Tujuan Pembelajaran
Di sini, kalian dapat mengetahui:
* Mengambil Value dari sebuah TextBox untuk:
  * Dioperasikan menjadi sebuah angka
  * Dioperasikan untuk bahan percobaan logika

### Listing
Langsung aja digas pol mang!

```vb
Private Sub Form_Load()
Text3.Enabled = False                 ' Sengaja, biar TextBoxnya ga bisa dijamah oleh tangan nakal bin iseng
Text1 = ""                            ' Tau lah gunanya "" untuk apa
Text2 = ""
Text3 = ""
End Sub

' Wohooo! Kali ini kita bakal belajar soal fungsi Val()
' Ok semisal kita mempunyai dua TextBox, anggaplah Text1 ama Text2 yang berisi
' Text1 = 3
' Text2 = 2
' Ya iya, kita tahu kalau mereka ditambah hasilnya menjadi 5
' Akan tetapi, bila kita menuliskan listingnya seperti:
' Output = Text1 + Text2
' Outputnya bukan menghasilkan 5 akan tetapi menghasilkan 32
' Mengapa demikian?
' Karena kita tidak mengambil value didalam Text1 dan Text2 untuk dijumlahkan
' Kalau cuma ngambil tanpa value, ini sama macam fungsi Concatenate di Excel (ngegabungin beberapa string menjadi satu string dalam sebuah cell)
' Jadi penulisan yang benar ialah
' Output = Val(Text1) + Val(Text2)

Private Sub Option1_Click()
Text3 = Val(Text1) > Val(Text2)
End Sub

Private Sub Option10_Click()
Text3 = Val(Text1) / Val(Text2)
End Sub

Private Sub Option11_Click()
Text3 = Val(Text1) ^ Val(Text2)
End Sub

Private Sub Option12_Click()
Text3 = Val(Text1) Mod Val(Text2)     ' Modular/Mod, untuk mengetahui sisa dari sebuah pembagian
End Sub

Private Sub Option2_Click()
Text3 = Val(Text1) < Val(Text2)
End Sub

Private Sub Option3_Click()
Text3 = Val(Text1) <> Val(Text2)
End Sub

Private Sub Option4_Click()
Text3 = Val(Text1) = Val(Text2)
End Sub

Private Sub Option5_Click()
Text3 = Val(Text1) >= Val(Text2)
End Sub

Private Sub Option6_Click()
Text3 = Val(Text1) <= Val(Text2)
End Sub

Private Sub Option7_Click()
Text3 = Val(Text1) + Val(Text2)
End Sub

Private Sub Option8_Click()
Text3 = Val(Text1) - Val(Text2)
End Sub

Private Sub Option9_Click()
Text3 = Val(Text1) * Val(Text2)
End Sub

```
