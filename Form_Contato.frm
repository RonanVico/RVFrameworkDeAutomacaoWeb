VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Contato 
   Caption         =   "                          Contato Ronan Vico"
   ClientHeight    =   1575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5595
   OleObjectBlob   =   "Form_Contato.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_Contato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub UserForm_Initialize()
  Mail.Value = "RonanVico@hotmail.com"
  Telefone.Value = "+55 11 94015 5925"
End Sub
