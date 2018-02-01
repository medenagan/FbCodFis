Attribute VB_Name = "FbCodFis"
'Implementazione Visual Basic per l'utilizzo della libreria FbCodFis.dll
'
'Agosto 2004
'Autore: Fabio Mereu

Option Explicit

'Dichiarazioni relative alle funzioni esportate
'***************************************** INIZIO *****************************************
Public Declare Sub AnalizzaCodiceFiscale Lib "FbCodFis.dll" (ByVal strCodice As String, ByVal Flags As Long, ByRef destAnalisi As CF_ANALISI_STRUCT)
Public Declare Function CreaCodiceFiscale Lib "FbCodFis.dll" (ByRef persData As CF_CODICE_FISCALE_STRUCT) As Long
Public Declare Function EstraiStrutturaDati Lib "FbCodFis.dll" (ByRef persData As CF_CODICE_FISCALE_STRUCT) As Long
Public Declare Function RecuperaCarattereControllo Lib "FbCodFis.dll" (ByVal STR_CF As String) As Long
Public Declare Sub RecuperaCodiceStandard Lib "FbCodFis.dll" (ByVal strOmoCodice As String, ByVal strCodiceDest As String, ByRef diffRet As CF_DIFFERENZIATO_STRUCT)
Public Declare Function RecuperaMessaggioErrore Lib "FbCodFis.dll" (ByVal Costante As Long, ByVal Tipo As Long, ByVal Destinazione As String, ByVal cchMax As Long) As Long
Public Declare Function SmistaCostanti Lib "FbCodFis.dll" (ByVal Costanti As Long, ByVal Tipo As Long, ByVal cllBack As Long, ByRef Extra As Any) As Boolean
'*****************************************  FINE  *****************************************

' Il prototipo della calling Back deve essere
' Sub MiaCallingBack(Costante As Long, Tipo As Long, PARAM As Long, ByRef Extra As Byte)

Public Const MAX_NOMINATIVO = 255
Public Const MAX_MSG_ERRORE = 255 + 1 'Grandezza massima dei messaggi di errore registrati nelle risorse stringa

'Dichiarazioni delle strutture
'CF_CODICE_FISCALE_STRUCT:
Public Type CF_CODICE_FISCALE_STRUCT
    CodiceFiscale(0 To 17) As Byte  ' char[16 + 2]
    '
    Cognome(0 To MAX_NOMINATIVO) As Byte ' char[MAX_NOMINATIVO + 1]
    Nome(0 To MAX_NOMINATIVO) As Byte ' char[MAX_NOMINATIVO + 1]
    '
    Sesso As Integer
    '
    Nascita_Anno As Integer
    Nascita_Mese As Integer
    Nascita_Giorno As Integer
    '
    Nascita_Luogo(0 To 5) As Byte    ' char[4 + 2]
End Type
'CF_ANALISI_STRUCT:
Public Type CF_ANALISI_STRUCT
    Analisi As Integer
    Totale As Integer
    Errore As Long
End Type
'CF_DIFFERENZIATO_STRUCT:
Public Type CF_DIFFERENZIATO_STRUCT
    Maschera As Integer
    totCarSost As Integer
    Errore As Long
End Type

' Dichiarazioni delle costanti statiche

'Definizione dei Mesi
Public Enum CF_MESE
    CF_MESE_GEN = 1
    CF_MESE_FEB = 2
    CF_MESE_MAR = 3
    CF_MESE_APR = 4
    CF_MESE_MAG = 5
    CF_MESE_GIU = 6
    CF_MESE_LUG = 7
    CF_MESE_AGO = 8
    CF_MESE_SET = 9
    CF_MESE_OTT = 10
    CF_MESE_NOV = 11
    CF_MESE_DIC = 12
End Enum

'Definizione del Sesso
Public Enum CF_SESSO
    CF_SESSO_F = 0
    CF_SESSO_M = 1
End Enum

'************************************ CF_ERRORE - INIZIO ************************************

'     NON_VALIDO: Strutturalmente incoerente; contiene caratteri
'                 INVALIDI - per esempio lettere al posto di
'                 numeri, caratteri speciali ()!?"£$<> ... invece
'                 di lettere - ; anche assenza - in particolare per
'                 il nome ed il cognome - di caratteri, oppure un
'                 numero insufficiente di questi.
'    FUORI_GAMMA: Strutturalmente valido, logicamente assurdo:
'                 è il caso del mese - che non può essere inferiore
'                 ad uno (gen) nè maggiore di 12 (dicembre) - ; del
'                 giorno - in relazione al mese e dell'anno se
'                 si tratta di febbraio - .

Public Enum CF_ERRORE
'NESSUN ERRORE
    CF_ERRORE_NO = &H0
    
' CARATTERI INVALIDI:
    
    'Nome e Cognome
    CF_ERRORE_COGNOME_NON_VALIDO = &H1
    CF_ERRORE_NOME_NON_VALIDO = &H2
    'Data: Giorno - Sesso, Mese, Anno
    CF_ERRORE_GIORNO_SESSO_NON_VALIDO = &H4
    CF_ERRORE_MESE_NON_VALIDO = &H8
    CF_ERRORE_ANNO_NON_VALIDO = &H10
    'Luogo
    CF_ERRORE_LUOGO_NON_VALIDO = &H20
    'Carattere di controllo
    CF_ERRORE_CF_CONTROLLO_NON_VALIDO = &H40

' ERRORI LOGICI:

    'Data: Giorno e Mese
    CF_ERRORE_GIORNO_FUORIGAMMA = &H80
    CF_ERRORE_MESE_FUORIGAMMA = &H100
    'Sesso
    CF_ERRORE_SESSO_FUORIGAMMA = &H200
    'Giorno\Sesso
    CF_ERRORE_GIORNO_SESSO_FUORIGAMMA = &H400

'BUFFER CODICE INSUFFICIENTE
    CF_ERRORE_BUFFER_CODICE_INSUFFICENTE = &H800
End Enum
'************************************ CF_ERRORE -  FINE  ************************************

' CF_ANALISI_FLAGS
Public Enum CF_ANALISI_FLAG
    CF_ANALISI_FLAG_CONTROLLO = &H1 'Obbliga AnalizzaCodiceFiscale ad accertarsi che il carattere di controllo sia valido anche dal punto di vista logico (usa GetCarattereControllo)
    CF_ANALISI_FLAG_DATA = &H2      'Consente ad AnalizzaCodiceFiscale di controllare la data da un punto di vista logico stretto: usa il mese ed eventualmente - se è Febbraio - l'anno per verificare il giorno. Se uno o entrambi i dati sono corrotti, li sostituisce: l'anno per default è bisestile; il mese per default è di 31 giorni. Se il flag è disattivato, non effettua alcun controllo e non segnala nessun errore relativo al giorno errato se questo non supera 31 (ovviamente anche femminile), indipendentemente dal mese e dall'anno.
End Enum

' CF_TIPO_COSTANTE
Public Enum CF_TIPOCONST
    CF_TIPOCONST_ERRORE = &H1000
    CF_TIPOCONST_MASCHERA = &H2000
End Enum

Function VirtualRecuperaMsgErrCompleto(Costanti As Long, Tipo As Long) As String
'Il prefisso 'Virtual' significa che la funzione
'non è presente nella libreria e costituisce un plus
'dell'implementazione.

'Questa funzione utilizza la procedura di libreria SmistaCostanti
'associata alla calling back di implementazione 'virtuale'
'VirtualCllBackErrorMessage e fornisce un messaggio di errore
'completo a partire da una combinazione biteriana di messaggi
'di errore singoli.
Dim strMessage As String
Call SmistaCostanti(Costanti, Tipo, AddressOf VirtualCllBackErrorMessage, strMessage)
VirtualRecuperaMsgErrCompleto = StrConv(strMessage, vbFromUnicode)
End Function

Private Sub VirtualCllBackErrorMessage(ByVal Costante As Long, ByVal Tipo As Long, _
        ByVal PARAM As Long, ByRef Extra As String)
        
'Il prefisso 'Virtual' significa che la funzione
'non è presente nella libreria e costituisce un plus
'dell'implementazione.

'Questa procedura di calling back recupera
'la stringa del messaggio di errore associato alla costante
'e compila la stringa 'Extra', creando un'unica stringa
'contenente tutti i messaggi di errore.
Dim tmpStr As String, pos As Long
tmpStr = Space(MAX_MSG_ERRORE)
Call RecuperaMessaggioErrore(Costante, Tipo, tmpStr, Len(tmpStr))
'Elimina l'eventuale vbNullChar
pos = InStr(tmpStr, vbNullChar)
If pos > 0 Then tmpStr = Left$(tmpStr, pos - 1)
'Accoda ad Extra il nuovo messaggio di errore
Extra = Extra & tmpStr & vbCrLf
End Sub
