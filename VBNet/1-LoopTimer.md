# Loop Timer
*Eh, ini tutorial VB.Net pertama ya? Ehehe, anyway, welcome :wink:*

![Example](https://cdn.discordapp.com/attachments/439384229483380746/447154287257190410/LoopTImer1.gif)

Ajaib tidak? Ok, kita akan membuatnya! Kalian hanya perlukan beberapa komponen seperti di bawah ini:

![Bahan](https://cdn.discordapp.com/attachments/439384229483380746/447154193489199115/unknown.png)

## Tujuan Pembelajaran
Di sini, kita akan mempelajari tentang:
* Timer
    * Single Tick
    * Looping
* Enable/Disable Button
* String Variable

### Listing
Ayo, dimari baca listingnya!

```vb
Public Class LoopTImer
    Dim SatuLabel As String             ' Variabel SatuLabel isinya harus String/Huruf

    ' Kondisi ketika Form dijalankan
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SatuLabel = "Apa"               ' Isi variabel SatuLabel dengan "Apa"
        Label1.Text = SatuLabel         ' Isi Label1 dengan isi dari variabel SatuLabel
        Label2.Text = ""                ' Kosongkan Label2
        Timer1.Interval = 1000          ' Beri Interval untuk Timer1 selama 1000 ms / 1 detik
    End Sub

    ' Kondisi ketika Button1 diklik
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Label1.Text = ""                ' Kosongkan Label1
        Label2.Text = SatuLabel         ' Isi Label2 dengan isi dari variabel SatuLabel
        Button1.Enabled = False         ' Buat Button1 menjadi tak bisa digunakan
        Timer1.Start()                  ' Jalankan Timer1
    End Sub

    ' Kondisi ketika trigger TImer1.Start() dikumandangkan
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ' Ok, kenapa gua langsung nulis kek gini dibawah?
        ' Interval = Selang
        ' Jadi, maksud dari Timer.Tick.Interval itu adalah:
        ' Berapa waktu yang dibutuhkan SEBELUM Timer bekerja
        ' Misalnya:
        ' Waktu lu ngeklik >>>> Interval >>>>> Aksi dari Timer
        ' Itu contoh kecilnya
        Label1.Text = SatuLabel         ' Isi Label1 dengan isi dari variabel SatuLabel
        Label2.Text = ""                ' Kosongkan Label2
        Button1.Enabled = True          ' Buat Button1 dapat digunakan
        Timer1.Stop()                   ' Hentikan Timer1
    End Sub
End Class
```

## Studi Kasus
Kan cuma single loop doang itu, **why 'bout some loopin?**

![Loopin!](https://cdn.discordapp.com/attachments/439384229483380746/447157164256002048/LoopTImer2.gif)

Yaudah, ini di bawah listingnya ehehe~

```vb
Public Class LoopTImer
    Dim SatuLabel As String             ' Variabel SatuLabel isinya harus String/Huruf

    ' Kondisi ketika Form dijalankan
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SatuLabel = "Apa"               ' Isi variabel SatuLabel dengan "Apa"
        Label1.Text = SatuLabel         ' Isi Label1 dengan isi dari variabel SatuLabel
        Label2.Text = ""                ' Kosongkan Label2
        Timer1.Interval = 1000          ' Beri Interval untuk Timer1 selama 1000 ms / 1 detik
        Timer2.Interval = 1000          ' Beri Interval untuk TImer2 selama 1000 ms / 1 detik
    End Sub

    ' Kondisi ketika Button1 diklik
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Button1.Text = "Switch" Then     ' Buat kondisi bilamana Button1 bertuliskan Switch
            Label1.Text = ""                ' Kosongkan Label1
            Label2.Text = SatuLabel         ' Isi Label2 dengan isi dari variabel SatuLabel
            Button1.Text = "Stop"
            Timer1.Start()                  ' Jalankan Timer1
        Else
            ' Hentikan Timer1 dan Timer2
            Timer1.Stop()
            Timer2.Stop()
            Button1.Text = "Switch"         ' Kembalikan tulisan Button1 menjadi Switch kembali
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label1.Text = SatuLabel         ' Isi Label1 dengan isi dari variabel SatuLabel
        Label2.Text = ""                ' Kosongkan Label2
        Timer1.Stop()                   ' Hentikan Timer1
        Timer2.Start()                  ' Jalankan Timer2
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Label1.Text = ""                ' Kosongkan Label1
        Label2.Text = SatuLabel         ' Isi Label2 dengan isi dari variabel SatuLabel
        Timer1.Start()                  ' Jalankan Timer1
        Timer2.Stop()                   ' Hentikan Timer2
    End Sub
End Class
```