Version =19
VersionRequired =19
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =55
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4320
    DatasheetFontHeight =10
    ItemSuffix =3
    Left =8820
    Top =1830
    Right =13140
    Bottom =5955
    DatasheetGridlinesColor =12632256
        0xcc680cc2fa67e240
    End
    Caption ="About"
    DatasheetFontName ="Arial"
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
            BackStyle =0
            FontName ="Tahoma"
        End
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
        End
            Height =4140
            BackColor =-2147483633
            Name ="Detail"
                    OverlapFlags =85
                    Left =120
                    Top =120
                    Width =4095
                    Height =3360
                    Name ="Label0"
                    Caption ="GPG Uage Demo Application\015\012Copyright (c) 2003 Brendan Kidwell\015\012\015\012"
                        "This Microsoft Access application demonstrates how to use the GNU Privacy Guard,"
                        " available for free at <http://www.gnupg.org/>, to send encrypted database updat"
                        "es to a central FTP server. Please see the file <readme.md> for usage notes. You"
                        " must follow the installation instructions in this file before the program will "
                        "work.\015\012\015\012Portions of this program were copied The Access Web <http:/"
                        "/www.mvps.org/access/>. In general, you may use the Visual Basic code found in t"
                        "his application however you wish, but be sure to read and respect any license in"
                        "formation you may find at the top of each module."
                End
                    Default = NotDefault
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =3720
                    Top =3660
                    Width =450
                    Height =405
                    Name ="cmdOK"
                    Caption ="OK"
                    OnClick ="[Event Procedure]"
                End
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =120
                    Top =3780
                    Width =1320
                    Height =255
                    ForeColor =1279872587
                    Name ="Label2"
                    Caption ="View readme.htm"
                    HyperlinkAddress ="readme.htm"
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdOK_Click()
On Error GoTo Err_cmdOK_Click


    DoCmd.Close

Exit_cmdOK_Click:
    Exit Sub

Err_cmdOK_Click:
    MsgBox Err.description
    Resume Exit_cmdOK_Click
    
End Sub
