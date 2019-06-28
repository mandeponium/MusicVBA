Attribute VB_Name = "Music"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Sub MortalKombat()
    
    Dim C0 As Long
    Dim Csharp0 As Long
    Dim D0 As Long
    Dim Dsharp0 As Long
    Dim E0 As Long
    Dim F0 As Long
    Dim Fsharp0 As Long
    Dim G0 As Long
    Dim Gsharp0 As Long
    Dim A0 As Long
    Dim Asharp0 As Long
    Dim B0 As Long
    Dim C1 As Long
    Dim Csharp1 As Long
    Dim D1 As Long
    Dim Dsharp1 As Long
    Dim E1 As Long
    Dim F1 As Long
    Dim Fsharp1 As Long
    Dim G1 As Long
    Dim Gsharp1 As Long
    Dim A1 As Long
    Dim Asharp1 As Long
    Dim B1 As Long
    Dim C2 As Long
    Dim Csharp2 As Long
    Dim D2 As Long
    Dim Dsharp2 As Long
    Dim E2 As Long
    Dim F2 As Long
    Dim Fsharp2 As Long
    Dim G2 As Long
    Dim Gsharp2 As Long
    Dim A2 As Long
    Dim Asharp2 As Long
    Dim B2 As Long
    Dim C3 As Long
    Dim Csharp3 As Long
    Dim D3 As Long
    Dim Dsharp3 As Long
    Dim D3 As Long
    Dim F3 As Long
    Dim Fsharp3 As Long
    Dim G3 As Long
    Dim Gsharp3 As Long
    Dim A3 As Long
    Dim Asharp3 As Long
    Dim B3 As Long
    Dim C4 As Long
    Dim Csharp4 As Long
    Dim D4 As Long
    Dim Dsharp4 As Long
    Dim E4 As Long
    Dim F4 As Long
    Dim Fsharp4 As Long
    Dim G4 As Long
    Dim Gsharp4 As Long
    Dim A4 As Long
    Dim Asharp4 As Long
    Dim B4 As Long
    Dim C5 As Long
    Dim Csharp5 As Long
    Dim D5 As Long
    Dim Dsharp5 As Long
    Dim E5 As Long
    Dim F5 As Long
    Dim Fsharp5 As Long
    Dim G5 As Long
    Dim Gsharp5 As Long
    Dim A5 As Long
    Dim Asharp5 As Long
    Dim B5 As Long
    Dim C6 As Long
    Dim Csharp6 As Long
    Dim D6 As Long
    Dim Dsharp6 As Long
    Dim E6 As Long
    Dim F6 As Long
    Dim Fsharp6 As Long
    Dim G6 As Long
    Dim Gsharp6 As Long
    Dim A6 As Long
    Dim Asharp6 As Long
    Dim B6 As Long
    Dim C7 As Long
    Dim Csharp7 As Long
    Dim D7 As Long
    Dim Dsharp7 As Long
    Dim E7 As Long
    Dim F7 As Long
    Dim Fsharp7 As Long
    Dim G7 As Long
    Dim Gsharp7 As Long
    Dim A7 As Long
    Dim Asharp7 As Long
    Dim B7 As Long
    Dim C8 As Long
    Dim Csharp8 As Long
    Dim D8 As Long
    Dim Dsharp8 As Long
    Dim E8 As Long
    Dim F8 As Long
    Dim Fsharp8 As Long
    Dim G8 As Long
    Dim Gsharp8 As Long
    Dim A8 As Long
    Dim Asharp8 As Long
    Dim B8 As Long


  
    C0 = 16.35          'C0
    Csharp0 = 17.32     'C#0/Db0
    D0 = 18.35          'D0
    Dsharp0 = 19.45     'D#
    E0 = 20.6           'E0
    F0 = 21.83          'F0
    Fsharp0 = 23.12     'F#0/Gb0 
    G0 = 24.5           'G0
    Gsharp0 = 25.96     'G#0/Ab0 
    A0 = 27.5           'A0
    Asharp0 = 29.14     'A#0/Bb0 
    B0 = 30.87          'B0
    C1 = 32.7           'C1
    Csharp1 = 34.65     'C#1/Db1 
    D1 = 36.71          'D1
    Dsharp1 = 38.89     'D#1/Eb1 
    E1 = 41.2           'E1
    F1 = 43.65          'F1
    Fsharp1 = 46.25     'F#1/Gb1 
    G1 = 49             'G1
    Gsharp1 = 51.91     'G#1/Ab1 
    A1 = 55             'A1
    Asharp1 = 58.27     'A#1/Bb1 
    B1 = 61.74          'B1
    C2 = 65.41          'C2
    Csharp2 = 69.3      'C#2/Db2 
    D2 = 73.42          'D2
    Dsharp2 = 77.78     'D#2/Eb2 
    E2 = 82.41          'E2
    F2 = 87.31          'F2
    Fsharp2 = 92.5      'F#2/Gb2 
    G2 = 98             'G2
    Gsharp2 = 103.83    'G#2/Ab2 
    A2 = 110            'A2
    Asharp2 = 116.54    'A#2/Bb2 
    B2 = 123.47         'B2
    C3 = 130.81         'C3
    Csharp3 = 138.59    'C#3/Db3 
    D3 = 146.83         'D3
    Dsharp3 = 155.56    'D#3/Eb3 
    D3 = 164.81         'E3
    F3 = 174.61         'F3
    Fsharp3 = 185       'F#3/Gb3 
    G3 = 196            'G3
    Gsharp3 = 207.65    'G#3/Ab3 
    A3 = 220      '     'A3
    Asharp3 = 233.08    'A#3/Bb3 
    B3 = 246.94         'B3
    C4 = 261.63         'C4
    Csharp4 = 277.18    'C#4/Db4 
    D4 = 293.66         'D4
    Dsharp4 = 311.13    'D#4/Eb4 
    E4 = 329.63         'E4
    F4 = 349.23         'F4
    Fsharp4 = 369.99    'F#4/Gb4 
    G4 = 392            'G4
    Gsharp4 = 415.3     'G#4/Ab4 
    A4 = 440            'A4
    Asharp4 = 466.16    'A#4/Bb4 
    B4 = 493.88         'B4
    C5 = 523.25         'C5
    Csharp5 = 554.37    'C#5/Db5 
    D5 = 587.33         'D5
    Dsharp5 = 622.25    'D#5/Eb5 
    E5 = 659.25         'E5
    F5 = 698.46         'F5
    Fsharp5 = 739.99    'F#5/Gb5 
    G5 = 783.99         'G5
    Gsharp5 = 830.61    'G#5/Ab5 
    A5 = 880            'A5
    Asharp5 = 932.33    'A#5/Bb5 
    B5 = 987.77         'B5
    C6 = 1046.5         'C6
    Csharp6 = 1108.73   ' C#6/Db6 
    D6 = 1174.66        'D6
    Dsharp6 = 1244.51   'D#6/Eb6 
    E6 = 1318.51        'E6
    F6 = 1396.91        'F6
    Fsharp6 = 1479.98   'F#6/Gb6 
    G6 = 1567.98        'G6
    Gsharp6 = 1661.22   'G#6/Ab6 
    A6 = 1760           'A6
    Asharp6 = 1864.66   'A#6/Bb6 
    B6 = 1975.53        'B6
    C7 = 2093           'C7
    Csharp7 = 2217.46   'C#7/Db7 
    D7 = 2349.32        'D7
    Dsharp7 = 2489.02   'D#7/Eb7 
    E7 = 2637.02        'E7
    F7 = 2793.83        'F7
    Fsharp7 = 2959.96   'F#7/Gb7 
    G7 = 3135.96        'G7
    Gsharp7 = 3322.44   'G#7/Ab7 
    A7 = 3520           'A7
    Asharp7 = 3729.31   'A#7/Bb7 
    B7 = 3951.07        'B7
    C8 = 4186.01        'C8
    Csharp8 = 4434.92   'C#8/Db8 
    D8 = 4698.63        'D8
    Dsharp8 = 4978.03   'D#8/Eb8 
    E8 = 5274.04        'E8
    F8 = 5587.65        'F8
    Fsharp8 = 5919.91   'F#8/Gb8 
    G8 = 6271.93        'G8
    Gsharp8 = 6644.88   'G#8/Ab8 
    A8 = 7040           'A8
    Asharp8 = 7458.62   'A#8/Bb8 
    B8 = 7902.13        'B8
    
    
    Dim BPM As Long
    BPM = 60

 
    Dim O As Long       'Whole note
    Dim O2 As Long      'Half note
    Dim O4 As Long      'Quarter note
    Dim O8 As Long      'Eigth note
    Dim O3_8 As Long    'Three eighth's note.
    Dim O16 As Long     'sixteenth note
    
       
    'The formula is 60 seconds / BPM * beat length * 1000 milliseconds
    
    O = (60 / BPM) * 4 * 1000
    O2 = (60 / BPM) * 2 * 1000
    O4 = (60 / BPM) * 1 * 1000
    O8 = (60 / BPM) * (1 / 8) * 1000
    O3_8 = (60 / BPM) * (3 / 8) * 1000
    O16 = (60 / BPM) * (1 / 16) * 1000



    Beep A4, O8
    Beep A4, O8
    Beep C5, O8
    Beep A4, O8
    Beep D5, O8
    Beep A4, O8
    Beep E5, O8
    Beep D5, O8
    
    
    Beep C5, O8
    Beep C5, O8
    Beep E5, O8
    Beep C5, O8
    Beep G5, O8
    Beep C5, O8
    Beep E5, O8
    Beep C5, O8
    
    
    Beep G4, O8
    Beep G4, O8
    Beep B4, O8
    Beep G4, O8
    Beep C5, O8
    Beep G4, O8
    Beep D5, O8
    Beep C5, O8
    
    
    Beep F4, O8
    Beep F4, O8
    Beep A4, O8
    Beep F4, O8
    Beep C5, O8
    Beep F4, O8
    Beep C5, O8
    Beep B4, O8
    
    '-----------
    
    Beep A4, O8
    Beep A4, O8
    Beep C5, O8
    Beep A4, O8
    Beep D5, O8
    Beep A4, O8
    Beep E5, O8
    Beep D5, O8
    

    Beep C5, O8
    Beep C5, O8
    Beep E5, O8
    Beep C5, O8
    Beep G5, O8
    Beep C5, O8
    Beep E5, O8
    Beep C5, O8
    
    
    Beep G4, O8
    Beep G4, O8
    Beep B4, O8
    Beep G4, O8
    Beep C5, O8
    Beep G4, O8
    Beep D5, O8
    Beep C5, O8
    
    
    Beep F4, O8
    Beep F4, O8
    Beep A4, O8
    Beep F4, O8
    Beep C5, O8
    Beep F4, O8
    Beep C5, O8
    Beep B4, O8
    
    
    '-------------
    
    
    Beep A4, O3_8
    Beep A4, O3_8
    Beep A4, O3_8
    Beep A4, O3_8
    Beep G4, O8
    Beep C5, O8
    
    Beep A4, O3_8
    Beep A4, O3_8
    Beep A4, O3_8
    Beep A4, O3_8
    Beep G4, O8
    Beep E4, O8
    
    
    Beep A4, O3_8
    Beep A4, O3_8
    Beep A4, O3_8
    Beep A4, O3_8
    Beep G4, O8
    Beep C5, O8
    
    Beep A4, O3_8
    Beep A4, O3_8
    Beep A4, O8
    Beep A4, O16
    Beep A4, O8
    Beep A4, O16
    Beep A4, O16
    
    Sleep O16
    Sleep O8
    
        
    '------------
        
    Beep A4, O3_8
    Beep A4, O3_8
    Beep A4, O3_8
    Beep A4, O3_8
    Beep G4, O8
    Beep C5, O8
    
    Beep A4, O3_8
    Beep A4, O3_8
    Beep A4, O3_8
    Beep A4, O3_8
    Beep G4, O8
    Beep E4, O8
    
    
    Beep A4, O3_8
    Beep A4, O3_8
    Beep A4, O3_8
    Beep A4, O3_8
    Beep G4, O8
    Beep C5, O8
    
    Beep A4, O3_8
    Beep A4, O3_8
    Beep A4, O8
    Beep A4, O16
    Beep A4, O8
    Beep A4, O16
    Beep A4, O16
    
    Sleep O16
    Sleep O8
    
    '----------
    
    
    Beep A4, O16
    Beep E5, O8
    Beep A4, O16
    Beep C5, O8
    Beep A4, O16
    Beep Asharp4, O8
    Beep A4, O16
    Beep C5, O8
    Beep Asharp4, O8
    Beep G5, O8
    
    Beep A4, O16
    Beep E5, O8
    Beep A4, O16
    Beep C5, O8
    Beep A4, O16
    Beep Asharp4, O8
    Beep A4, O16
    Beep C5, O8
    Beep Asharp4, O8
    Beep G5, O8
    
    Beep A4, O16
    Beep E5, O8
    Beep A4, O16
    Beep C5, O8
    Beep A4, O16
    Beep Asharp4, O8
    Beep A4, O16
    Beep C5, O8
    Beep Asharp4, O8
    Beep G5, O8
    
    '--------------
    
    Beep A5, O3_8
    Beep E5, O3_8
    Beep D5, O3_8
    
    Sleep O16
    
    Beep Asharp5, O8
    Beep A5, O4

    Beep A5, O3_8
    Beep E5, O3_8
    Beep D5, O3_8
    
    Sleep O16
    
    Beep Asharp5, O8
    Beep A5, O4
    
    
End Sub
