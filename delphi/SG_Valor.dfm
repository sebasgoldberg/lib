object FormValor: TFormValor
  Left = 218
  Top = 140
  Width = 296
  Height = 113
  BorderIcons = []
  Caption = 'Nuevo valor'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnActivate = FormActivate
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBox1: TGroupBox
    Left = 8
    Top = 0
    Width = 265
    Height = 41
    Anchors = [akLeft, akTop, akRight, akBottom]
    TabOrder = 0
    object Label1: TLabel
      Left = 8
      Top = 16
      Width = 24
      Height = 13
      Caption = 'Valor'
    end
    object EditValor: TEdit
      Left = 136
      Top = 12
      Width = 121
      Height = 21
      TabOrder = 0
    end
  end
  object BitBtnAceptar: TBitBtn
    Left = 84
    Top = 46
    Width = 89
    Height = 25
    Anchors = [akRight, akBottom]
    Caption = 'Aceptar'
    TabOrder = 1
    OnClick = BitBtnAceptarClick
  end
  object BitBtnCancelar: TBitBtn
    Left = 183
    Top = 46
    Width = 89
    Height = 25
    Anchors = [akRight, akBottom]
    Caption = 'Cancelar'
    TabOrder = 2
    OnClick = BitBtnCancelarClick
  end
end
