Attribute VB_Name = "ModMain"
Option Explicit
'============================================================================
' ModMain.bas - Module shell
'
' This module has been refactored. All code has been migrated to:
'   Constants.bas  - All Const definitions
'   Models.bas     - Type definitions and global data arrays
'   FileParser.bas - Low-level file I/O (GetWord, PutWord, etc.)
'   CsvLoader.bas  - CSV localization file loading (unified)
'   FileLoader.bas - Data file loading (Load*File, Read*Line)
'   FileSaver.bas  - Data file saving (Save*File, Output*Line)
'   Init.bas       - Initialization (PublicInit, InitWarbandInfo, etc.)
'   Utils.bas      - Utility functions (Swap*, Is*, etc.)
'
' This shell is kept for backward compatibility with the VBP project file.
'============================================================================
