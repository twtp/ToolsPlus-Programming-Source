Attribute VB_Name = "GlobalsPoInv"
'---------------------------------------------------------------------------------------
' Module    : GlobalsPoInv
' DateTime  : 10/27/2005 13:29
' Author    : briandonorfio
' Purpose   : Define globals for the po/inv program.
'---------------------------------------------------------------------------------------

Option Explicit

'******************************************************************************
'  CONSTANTS
'******************************************************************************

Public Const CONF_FILE As String = "s:\mastest\mas90-signs\conf_files\PoInv.conf"

Public Const DISTRIBUTION_DIR As String = "s:\mastest\mas90-signs\A_Dist\"
Public Const REPORTS_DB As String = "reports.mdb"
Public Const SIGNS_DB As String = "print_signs.mdb"
Public Const PO_DB As String = "purchase_orders.mdb"
Public Const MISC_DB As String = "misc_signs.mdb"

Public Const MUTEX_INV_CHANGE As String = "s:\mastest\mas90-signs\export\export_invchange_complete.txt"

'Public Const USE_ALT_DATABASE As Boolean = True


'******************************************************************************
' DATABASE HANDLES
'******************************************************************************
Public DB As DBConn.DatabaseSingleton      'handle to MSSQL
Public MASDB As DBConn.DatabaseSingleton   'handle to MAS200
'Public UPCDB As DBConn.DatabaseSingleton   'handle to BARCODE
'Public SHIPDB As DBConn.DatabaseSingleton  'handle to SHIPPING




'******************************************************************************
'  GLOBAL VARIABLES
'******************************************************************************

Public CUR_POS_OPT As String        'where the cursor should go on entry to a field
                                    '  "start"   - at the beginning of the field
                                    '  "end"     - at the end of the field
                                    '  "select"  - highlight entire field
                                    '  "default" - default, positions where you click

Public NotifyChanges As PerlArray   'array to save item changes, actually an array of arrays.
                                    'each change should push on a variant array as follows:
                                    '  NotifyChanges(i)(0) = itemnumber
                                    '  NotifyChanges(i)(1) = change type
                                    '    "discontinued"
                                    '    "store price"
                                    '    "web price"
                                    '    "cost"
                                    '  NotifyChanges(i)(2) = old value
                                    'everything should be handled by a wrapper function,
                                    'UpdateNotificationEmail.

Public WARN_ON_IMPORT As Boolean    'whether we should warn other users that an import or
                                    'export to/from mas200 is currently running.

Public FBF_FIELDS(7) As String      'for saving filtering by form, must be changed if the
Public FBF_TEXT(7) As String        'number of spaces for things increase
Public FBF_ANDOR(6) As String

Public FBS_PATHS(1 To 5) As String  'for saving paths in a filter by search.

'Public POINV_SESSION_ID As Long     'login session number

Public ImageCache As Dictionary     'items and sections, binary image indexed by file path

Public CURRENT_WEBSITE_ID As Long   'selected website id, will be set to 0 (tp) initially

Public META_NOINDEX_MOLLY_ON As Boolean 'on/off the molly guard for signs and inv no-index option

Public INV_MAINT_FOLLIES_EDIT As Boolean 'allow editing of follies state
