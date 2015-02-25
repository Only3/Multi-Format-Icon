# Multi-Format-Icon
Extract Or Create Multi Format Icon

# Create

Dim I As New MultiIcon({"1.ico", "2.ico"})

I.SaveMultiIcon("New.ico")

# Extract And Save

Dim I As New MultiIcon("Multi.ico")

'Icon = I.Icons(0)

I.SaveIcon(0, "0.ico")

If I.Count > 1 Then I.SaveIcons("...\Desktop")
