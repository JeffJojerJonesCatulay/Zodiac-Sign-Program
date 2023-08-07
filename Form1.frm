VERSION 5.00
Begin VB.Form Exercise2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zodiac Sign"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4.365
   ScaleMode       =   5  'Inch
   ScaleWidth      =   8.438
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR"
      Height          =   315
      Left            =   6120
      TabIndex        =   6
      Top             =   5880
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Form1.frx":483E5
      Left            =   9000
      List            =   "Form1.frx":48428
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   5400
      Width           =   1935
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Form1.frx":484AA
      Left            =   5880
      List            =   "Form1.frx":4850B
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   5400
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Form1.frx":48582
      Left            =   2640
      List            =   "Form1.frx":485AA
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FIND"
      Height          =   315
      Index           =   0
      Left            =   4320
      TabIndex        =   0
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "<<<Go Back"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   360
      TabIndex        =   7
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3975
      Left            =   8640
      TabIndex        =   5
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   4335
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   960
      Width           =   2655
   End
End
Attribute VB_Name = "Exercise2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_Click()

End Sub

Private Sub Command1_Click(Index As Integer)
Dim month As String
month = Combo2.Text
Dim day As Integer
day = Combo3.Text
Dim year As Integer
year = Combo1.Text

'Year
If (year - 2000) Mod 12 = 0 Then
    Label2.Caption = "Year of the Dragon" + vbNewLine + vbNewLine + "A powerful sign, those born under the Chinese Zodiac sign of the Dragon are energetic and warm-hearted, charismatic, lucky at love and egotistic. They're natural born leaders, good at giving orders and doing what's necessary to remain on top. Compatible with Monkey and Rat."
ElseIf (year - 2000) Mod 12 = 1 Then
    Label2.Caption = "Year of the Snake" + vbNewLine + vbNewLine + "Those born under the Chinese Zodiac sign of the Snake are seductive, gregarious, introverted, generous, charming, good with money, analytical, insecure, jealous, slightly dangerous, smart, they rely on gut feelings, are hard-working and intelligent. Compatible with Rooster or Ox."
ElseIf (year - 2000) Mod 12 = 2 Then
    Label2.Caption = "Year of the Horse" + vbNewLine + vbNewLine + "Those born under the Chinese Zodiac sign of the Horse love to roam free. They're energetic, self-reliant, money-wise, and they enjoy traveling, love and intimacy. They're great at seducing, sharp-witted, impatient and sometimes seen as a drifter. Compatible with Dog or Tiger."
ElseIf (year - 2000) Mod 12 = 3 Then
    Label2.Caption = "Year of the Goat" + vbNewLine + vbNewLine + "Those born under the Chinese Zodiac sign of the Goat enjoy being alone in their thoughts. They're creative, thinkers, wanderers, unorganized, high-strung and insecure, and can be anxiety-ridden. They need lots of love, support and reassurance. Appearance is important too. Compatible with Pig or Rabbit."
ElseIf (year - 2000) Mod 12 = 4 Then
    Label2.Caption = "Year of the Monkey" + vbNewLine + vbNewLine + "Those born under the Chinese Zodiac sign of the Monkey thrive on having fun. They're energetic, upbeat, and good at listening but lack self-control. They like being active and stimulated and enjoy pleasing self before pleasing others. They're heart-breakers, not good at long-term relationships, morals are weak. Compatible with Rat or Dragon."
ElseIf (year - 2000) Mod 12 = 5 Then
    Label2.Caption = "Year of the Rooster" + vbNewLine + vbNewLine + "Those born under the Chinese Zodiac sign of the Rooster are practical, resourceful, observant, analytical, straightforward, trusting, honest, perfectionists, neat and conservative. Compatible with Ox or Snake."
ElseIf (year - 2000) Mod 12 = 6 Then
    Label2.Caption = "Year of the Dog" + vbNewLine + vbNewLine + "Those born under the Chinese Zodiac sign of the Dog are loyal, faithful, honest, distrustful, often guilty of telling white lies, temperamental, prone to mood swings, dogmatic, and sensitive. Dogs excel in business but have trouble finding mates. Compatible with Tiger or Horse."
ElseIf (year - 2000) Mod 12 = 7 Then
    Label2.Caption = "Year of the Pig" + vbNewLine + vbNewLine + "Those born under the Chinese Zodiac sign of the Pig are extremely nice, good-mannered and tasteful. They're perfectionists who enjoy finer things but are not perceived as snobs. They enjoy helping others and are good companions until someone close crosses them, then look out! They're intelligent, always seeking more knowledge, and exclusive. Compatible with Rabbit or Goat."
ElseIf (year - 2000) Mod 12 = 8 Then
    Label2.Caption = "Year of the Rat" + vbNewLine + vbNewLine + "Those born under the Chinese Zodiac sign of the Rat are quick-witted, clever, charming, sharp and funny. They have excellent taste, are a good friend and are generous and loyal to others considered part of its pack. Motivated by money, can be greedy, is ever curious, seeks knowledge and welcomes challenges. Compatible with Dragon or Monkey."
ElseIf (year - 2000) Mod 12 = 9 Then
    Label2.Caption = "Year of the Ox" + vbNewLine + vbNewLine + "Another of the powerful Chinese Zodiac signs, the Ox is steadfast, solid, a goal-oriented leader, detail-oriented, hard-working, stubborn, serious and introverted but can feel lonely and insecure. Takes comfort in friends and family and is a reliable, protective and strong companion. Compatible with Snake or Rooster."
ElseIf (year - 2000) Mod 12 = 10 Then
    Label2.Caption = "Year of the Tiger" + vbNewLine + vbNewLine + "Those born under the Chinese Zodiac sign of the Tiger are authoritative, self-possessed, have strong leadership qualities, are charming, ambitious, courageous, warm-hearted, highly seductive, moody, intense, and they're ready to pounce at any time. Compatible with Horse or Dog."
Else
    Label2.Caption = "Year of the Rabbit" + vbNewLine + vbNewLine + "Those born under the Chinese Zodiac sign of the Rabbit enjoy being surrounded by family and friends. They're popular, compassionate, sincere, and they like to avoid conflict and are sometimes seen as pushovers. Rabbits enjoy home and entertaining at home. Compatible with Goat or Pig."


End If

'Aquarius
If month = "January" Then
    If day >= 20 Then
        Label1(0).Caption = "Aquarius" + vbNewLine + "The Water Bearer" + vbNewLine + "(Jan 20 - Feb 18)" + vbNewLine + vbNewLine + "The mad Scientist and humanitarian of the horoscope wheel, futuristic. Aquariuss energey helps us innovate and unite for social justice."
'Capricorn
    ElseIf day <= 19 Then
        Label1(0).Caption = "Capricorn" + vbNewLine + "The Goat" + vbNewLine + "(Dec 20 - Jan 19)" + vbNewLine + vbNewLine + "The measured master planner of the horoscope family, Capricorn energy teaches us the power of structure and long-term goals"
    End If

ElseIf month = "February" Then
    If day <= 18 Then
        Label1(0).Caption = "Aquarius" + vbNewLine + "The Water Bearer" + vbNewLine + "(Jan 20 - Feb 18)" + vbNewLine + vbNewLine + "The mad Scientist and humanitarian of the horoscope wheel, futuristic. Aquariuss energey helps us innovate and unite for social justice."
'Pisces
    ElseIf day <= 28 Then
        Label1(0).Caption = "Pisces" + vbNewLine + "The Fish" + vbNewLine + "(Feb 19 - Mar 20)" + vbNewLine + vbNewLine + "The dreamer and healer of the horoscope family, Picses energy awakens compassion, imaggination and artistry, uniting us as one."
    End If


ElseIf month = "March" Then
    If day <= 20 Then
        Label1(0).Caption = "Pisces" + vbNewLine + "The Fish" + vbNewLine + "(Feb 19 - Mar 20)" + vbNewLine + vbNewLine + "The dreamer and healer of the horoscope family, Picses energy awakens compassion, imaggination and artistry, uniting us as one."
'Aries
    ElseIf day >= 21 Then
        Label1(0).Caption = "Aries" + vbNewLine + "The Ram" + vbNewLine + "(Mar 21 - Apr 19)" + vbNewLine + vbNewLine + "The pioneer and trailblazer of the horoscope wheel, Aries energy helps us initiate, fight for our beliefs and fearlessly put ourselves out there."
    End If


ElseIf month = "April" Then
    If day <= 19 Then
        Label1(0).Caption = "Aries" + vbNewLine + "The Ram" + vbNewLine + "(Mar 21 - Apr 19)" + vbNewLine + vbNewLine + "The pioneer and trailblazer of the horoscope wheel, Aries energy helps us initiate, fight for our beliefs and fearlessly put ourselves out there."
'Taurus
    ElseIf day >= 20 Then
        Label1(0).Caption = "Taurus" + vbNewLine + "The Bull" + vbNewLine + "(Apr 20 - May 20)" + vbNewLine + vbNewLine + "The perrsistent provider of the horoscope family, Taurus energy helps us seek security, enjoy earthly pleasures and get the job done."
    End If


ElseIf month = "May" Then
    If day <= 20 Then
        Label1(0).Caption = "Taurus" + vbNewLine + "The Bull" + vbNewLine + "(Apr 20 - May 20)" + vbNewLine + vbNewLine + "The perrsistent provider of the horoscope family, Taurus energy helps us seek security, enjoy earthly pleasures and get the job done."
'Gemini
    ElseIf day >= 21 Then
        Label1(0).Caption = "Gemini" + vbNewLine + "The Twins" + vbNewLine + "(May 21 - Jun 20)" + vbNewLine + vbNewLine + "The most versatile and vibrant horoscope sign, Gemini energy helps us communicate collaborate and fly our freak flags at full mast."
    End If


ElseIf month = "June" Then
    If day <= 20 Then
        Label1(0).Caption = "Gemini" + vbNewLine + "The Twins" + vbNewLine + "(May 21 - Jun 20)" + vbNewLine + vbNewLine + "The most versatile and vibrant horoscope sign, Gemini energy helps us communicate collaborate and fly our freak flags at full mast."
'Cancer
    ElseIf day >= 21 Then
        Label1(0).Caption = "Cancer" + vbNewLine + "The Crab" + vbNewLine + "(June 21 - July 22)" + vbNewLine + vbNewLine + "The natural nurturer of the horoscope wheel, Cancer energy helps us connect with our feelings, plant deep roots and feather our family nests."
    End If


ElseIf month = "July" Then
    If day <= 22 Then
        Label1(0).Caption = "Cancer" + vbNewLine + "The Crab" + vbNewLine + "(June 21 - July 22)" + vbNewLine + vbNewLine + "The natural nurturer of the horoscope wheel, Cancer energy helps us connect with our feelings, plant deep roots and feather our family nests."
'Leo
    ElseIf day >= 23 Then
        Label1(0).Caption = "Leo" + vbNewLine + "The Lion" + vbNewLine + "(July 23 - Aug 22)" + vbNewLine + vbNewLine + "The drama queen and regal ruler of the horoscope clan, Leo energy helps us shine, express ourselves boldly and wear our hearts on our sleeves."
    End If


ElseIf month = "August" Then
    If day <= 22 Then
        Label1(0).Caption = "Leo" + vbNewLine + "The Lion" + vbNewLine + "(July 23 - Aug 22)" + vbNewLine + vbNewLine + "The drama queen and regal ruler of the horoscope clan, Leo energy helps us shine, express ourselves boldly and wear our hearts on our sleeves."
'Virgo
    ElseIf day >= 23 Then
        Label1(0).Caption = "Virgo" + vbNewLine + "The Virgin" + vbNewLine + "(Aug 23 - Sep 22)" + vbNewLine + vbNewLine + "The masterful helper of the hotoscope wheel, Virgo Energy teaches us to serve, do impeccable work and prioritize wellbeing of ourselves, our loved ones and he planet."
    End If


ElseIf month = "September" Then
    If day <= 22 Then
        Label1(0).Caption = "Virgo" + vbNewLine + "The Virgin" + vbNewLine + "(Aug 23 - Sep 22)" + vbNewLine + vbNewLine + "The masterful helper of the hotoscope wheel, Virgo Energy teaches us to serve, do impeccable work and prioritize wellbeing of ourselves, our loved ones and he planet."
'Libra
    ElseIf day >= 23 Then
        Label1(0).Caption = "Libra" + vbNewLine + "The Scales" + vbNewLine + "(Sep 23 - Oct 22)" + vbNewLine + vbNewLine + "The balanced beautifier of the horoscope family, Libra energy inspires us to seek peace, harmony and cooperation and to do it with style and grace."
    End If


ElseIf month = "October" Then
    If day <= 22 Then
        Label1(0).Caption = "Libra" + vbNewLine + "The Scales" + vbNewLine + "(Sep 23 - Oct 22)" + vbNewLine + vbNewLine + "The balanced beautifier of the horoscope family, Libra energy inspires us to seek peace, harmony and cooperation and to do it with style and grace."
'Scorpio
    ElseIf day >= 23 Then
        Label1(0).Caption = "Scorpion" + vbNewLine + "The Scorpion" + vbNewLine + "(Oct 23 - Nov 21)" + vbNewLine + vbNewLine + "The most intense and focused of the horoscope signs, Scorpio energy helps us dive deep, merge our super powers and form bonds that are built to last."
    End If


ElseIf month = "November" Then
    If day <= 21 Then
        Label1(0).Caption = "Scorpion" + vbNewLine + "The Scorpion" + vbNewLine + "(Oct 23 - Nov 21)" + vbNewLine + vbNewLine + "The most intense and focused of the horoscope signs, Scorpio energy helps us dive deep, merge our super powers and form bonds that are built to last"
'Sagittarius
    ElseIf day >= 22 Then
        Label1(0).Caption = "Sagittarius" + vbNewLine + "The Archer" + vbNewLine + "(Nov 22 - Dec 21)" + vbNewLine + vbNewLine + "The wordly adventurer of the horoscope wheel, Sagittarius energy inspire us to dream big, chase the impossible and take fearless risk"
    End If


ElseIf month = "December" Then
    If day <= 21 Then
        Label1(0).Caption = "Sagittarius" + vbNewLine + "The Archer" + vbNewLine + "(Nov 22 - Dec 21)" + vbNewLine + vbNewLine + "The wordly adventurer of the horoscope wheel, Sagittarius energy inspire us to dream big, chase the impossible and take fearless risk"
'Capricorn
    ElseIf day >= 22 Then
        Label1(0).Caption = "Capricorn" + vbNewLine + "The Goat" + vbNewLine + "(Dec 20 - Jan 19)" + vbNewLine + vbNewLine + "The measured master planner of the horoscope family, Capricorn energy teaches us the power of structure and long-term goals"
    End If
End If

End Sub

Private Sub Command2_Click()
Label1(0).Caption = ""
Label2.Caption = ""

End Sub

Private Sub Label3_Click()
Exercise2.Hide
MainForm.Show
End Sub
