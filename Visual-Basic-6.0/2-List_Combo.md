# ListBox dan ComboBox

Ok, kali ini kita akan belajar untuk mengoperasikan ListBox dan ComboBox dengan sederhana.

![Contoh Form](https://media.discordapp.net/attachments/439384229483380746/439384536112037889/VB6_2018-04-27_19-17-12.png)

Nah kira-kira seperti itu Form yang akan kita buat. Nah, untuk simulasinya, kita akan memasukkan nama-nama guru yang berada di dalam ComboBox ke dalam ListBox.

![Isi ComboBox](https://media.discordapp.net/attachments/439384229483380746/439387195057504256/VB6_2018-04-27_19-26-41.png)

Berikut di bawah adalah listingnya:

```vb
Private Sub Form_Load()

' Mari kita berkenalan dengan fungsi Combo.AddItem
' Fungsi AddItem pada ComboBox yaitu memasukkan sebuah String/Kata kedalam ComboBox
' Misalnya Combo1.AddItem "Ikram"
' Berarti ketika Form dijalankan, akan secara otomatis nama Ikram akan muncul di dalam Combo1

Combo1.AddItem "Irman Aras"
Combo1.AddItem "Andy Rafi"
Combo1.AddItem "Achmad"
Combo1.AddItem "Marlina"
Combo1.AddItem "Sumarni"
End Sub

Private Sub Command1_Click()
List1.AddItem Combo1.Text             ' List1 akan menambah item berupa string dari pilihan Combo1 yang Client pilih
Combo1.SetFocus                       ' Buat Combo1 menjadi focus lagi
End Sub

Private Sub Command2_Click()
List1.RemoveItem List1.ListIndex      ' List1 akan menghapus item yang Client pilih
End Sub

Private Sub Command3_Click()
List1.Clear                           ' List1 akan mengosongkan keseluruhan item
Combo1.SetFocus
End Sub
```
