Version =19
VersionRequired =19
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5700
    DatasheetFontHeight =10
    ItemSuffix =24
    Left =3675
    Top =1590
    Right =9660
    Bottom =5880
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x5f6d5044fb67e240
    End
    RecordSource ="addresses"
    Caption ="Addresses"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ListBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin Tab
            BackStyle =0
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =4296
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =120
                    Width =2568
                    ColumnWidth =2568
                    Name ="LastName"
                    ControlSource ="LastName"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =1380
                            Height =240
                            Name ="LastName_Label"
                            Caption ="Last Name"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =480
                    Width =2568
                    ColumnWidth =2568
                    TabIndex =1
                    Name ="FirstName"
                    ControlSource ="FirstName"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =480
                            Width =1380
                            Height =240
                            Name ="FirstName_Label"
                            Caption ="First Name"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =840
                    Width =2568
                    ColumnWidth =2568
                    TabIndex =2
                    Name ="Street"
                    ControlSource ="Street"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =840
                            Width =1380
                            Height =240
                            Name ="Street_Label"
                            Caption ="Street"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =1200
                    Width =2568
                    ColumnWidth =2568
                    TabIndex =3
                    Name ="City"
                    ControlSource ="City"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1200
                            Width =1380
                            Height =240
                            Name ="City_Label"
                            Caption ="City"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =1560
                    Width =384
                    ColumnWidth =384
                    TabIndex =4
                    Name ="State"
                    ControlSource ="State"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1560
                            Width =1380
                            Height =240
                            Name ="State_Label"
                            Caption ="State"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3060
                    Top =1560
                    Width =972
                    ColumnWidth =972
                    TabIndex =5
                    Name ="Zip"
                    ControlSource ="Zip"
                    InputMask ="00000C####"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2700
                            Top =1560
                            Width =288
                            Height =228
                            Name ="Zip_Label"
                            Caption ="Zip"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =1920
                    Width =2568
                    ColumnWidth =2568
                    TabIndex =6
                    Name ="Email"
                    ControlSource ="Email"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1920
                            Width =1380
                            Height =240
                            Name ="Email_Label"
                            Caption ="Email Address"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =2280
                    Width =1728
                    ColumnWidth =1728
                    TabIndex =7
                    Name ="Phone"
                    ControlSource ="Phone"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2280
                            Width =1380
                            Height =240
                            Name ="Phone_Label"
                            Caption ="Phone Number"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =2640
                    Width =1140
                    ColumnWidth =1140
                    TabIndex =8
                    Name ="Birthday"
                    ControlSource ="Birthday"
                    Format ="Short Date"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2640
                            Width =1380
                            Height =240
                            Name ="MemberSince_Label"
                            Caption ="Birthday"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =3000
                    Width =4020
                    Height =816
                    ColumnWidth =3000
                    TabIndex =10
                    Name ="Comments"
                    ControlSource ="Comments"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3000
                            Width =1380
                            Height =240
                            Name ="Comments_Label"
                            Caption ="Comments"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2760
                    Top =2640
                    Width =246
                    Height =246
                    TabIndex =9
                    Name ="cmdBirthdayCal"
                    OnClick ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada000000000000000d0fffffffffffff0a ,
                        0x0f7777777fffff0d0f7f7f7f7fffff0a0f77777777777f0d0f7f7f7f7f7f7f0a ,
                        0x0f77777777777f0d0f7f7f7f7f7f7f0a0f77777777777f0d0f7f7f7f7f7f7f0a ,
                        0x0f77777777777f0d0fffffffffffff0a0f777777fff77f0d0fffffffffffff0a ,
                        0x000000000000000d000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1500
                    Top =3960
                    Width =1740
                    TabIndex =11
                    Name ="RecordCreated"
                    ControlSource ="RecordCreated"
                    Format ="General Date"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =3960
                            Width =1248
                            Height =228
                            Name ="Label23"
                            Caption ="Record Created:"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdBirthdayCal_Enter()
InputDateField Birthday
End Sub
