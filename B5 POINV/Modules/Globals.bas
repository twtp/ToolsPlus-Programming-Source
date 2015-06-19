Attribute VB_Name = "Globals"
'---------------------------------------------------------------------------------------
' Module    : Globals
' DateTime  : 8/19/2005 16:29
' Author    : briandonorfio
' Purpose   : Define overarching constants and global variables, should be the same for
'             every toolsplus app.
'
'             Dependencies:
'               - none, but this should be considered a dependency for everything else.
'---------------------------------------------------------------------------------------

Option Explicit


'******************************************************************************
'  DEBUG MODE -- DON'T TOUCH THIS
'******************************************************************************
Public Const DEBUG_MODE As Boolean = False


'******************************************************************************
'  API FUNCTIONS
'******************************************************************************
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


'******************************************************************************
'  STRUCTS and ENUMERATED TYPES
'******************************************************************************
Public Type outlookContact
    fullName As String
    jobTitle As String
    Address As String
    City As String
    State As String
    ZipCode As String
    phoneNumber1 As String
    phoneExtension1 As String
    phoneNumber2 As String
    phoneExtension2 As String
    cellPhoneNumber As String
    faxNumber As String
    emailAddress As String
    webAddress As String
    body As String
End Type

Public Enum ErrorCodes
    CombinatoricsInvalidArguments = 1
    CombinatoricsArgumentTypeMismatch
    TransliterateInvalidArguments
    GenericMBArgumentTypeMismatch
End Enum



'******************************************************************************
'  CONSTANTS
'******************************************************************************
Public Const MODAL As Long = 1

Public Const SIGN_MASTERSIGN_ID As Long = 22

Public Const PERL As String = "\\toolsplus04\perl\bin\perl.exe"
Public Const WPERL As String = "\\toolsplus04\perl\bin\wperl.exe"

Public Const FREIGHTQUOTE_SINGLE As String = "s:\mastest\mas90-signs\freight_quote\poinv_freight_quote.pl"

Public Const WPC As String = "s:\mastest\mas90-signs\wpc_new\wpc_basic.pl"

'Public Const CL_EMAIL As String = "s:\mastest\mas90-signs\command_line_email.pl"
Public Const CL_EMAIL As String = "s:\mastest\mas90-signs\command_line_email.bat"

Public Const SHOPPING_FEEDS As String = "s:\mastest\mas90-signs\update_shopping_engine_feeds.pl"
'Public Const AFFILIATE_FEED As String = "s:\mastest\mas90-signs\upload_affiliate_feed.pl"
Public Const REFRESH_VENDOR_QTYS As String = "s:\mastest\mas90-signs\refresh_vendor_qtys.pl"

Public Const KWD_Y_OFFLOAD As String = "s:\mastest\mas90-signs\keywords_scripts\generate_overture_upload.pl"
Public Const KWD_Y_IMP_PARSE As String = "s:\mastest\mas90-signs\keywords_scripts\parse_yahoo_import.pl"
Public Const KWD_O_IMP_PARSE As String = "s:\mastest\mas90-signs\keywords_scripts\parse_overture_import.pl"
Public Const KWD_O_CLEAR_FLAGS As String = "s:\mastest\mas90-signs\keywords_scripts\sync_overture_database.pl"

Public Const FQ_STANDALONE_APP = "s:\mastest\mas90-signs\FreightCalc.exe"
Public Const FQ_GUI_APP = "s:\mastest\mas90-signs\freight_quote\fq_gui.pl"
Public Const WPC_GUI_APP = "s:\mastest\mas90-signs\wpc_new\wpc_gui.pl"
Public Const SIGNS_APP = "s:\mastest\mas90-signs\signs\signs.pl"
Public Const AR_PLAYLIST_MAKER = "h:\audio runner\programs\PlaylistMaker.exe"
'Public Const SHIPDB_UPD_PROCESS = "s:\mastest\mas90-signs\A_Dist\ShipDBUpdater.exe"

Public Const ORDER_PROCESSING_EMAIL = "orderprocessing@tools-plus.com"

'Public Const SHIPPING_GUY_PHONE_1 As String = "203-695-3424"
Public Const SHIPPING_GUY_PHONE_1 As String = ""
Public Const SHIPPING_GUY_PHONE_2 As String = "203-573-0750 ext 823"

Public Const COLOR_BUTTON_TEXT As Long = &H80000012
Public Const WHITE As Long = &HFFFFFF
Public Const LT_GREY As Long = &HE0E0E0
Public Const YELLOW As Long = &HFFFF&
Public Const BLACK As Long = &H80000008
Public Const BLUE As Long = &HFF0000
Public Const RED As Long = &HFF&
Public Const GREEN As Long = &HFF00&
Public Const DARK_GREEN As Long = &H8000&
Public Const BG_GREY As Long = &H8000000F
Public Const LT_BLUE As Long = &HFFBB00

'******************************************************************************
'  GLOBAL VARIABLES
'******************************************************************************
Public Mouse As MouseHourglass        'semaphore based mousepointer control

Public IMPORT_GOING As Boolean        'whether we know if an import/export is going
Public EXPORT_GOING As Boolean

Public DESTINATION_DIR As String      'temp dir...should be a constant, but uses Environ()

Public USING_SECURITY As Boolean    'enable/disable user-level security
