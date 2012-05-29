#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=pictographs-fire_extinguisher_inv.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;This AutoIt wrapper is only used to provide an EXE with a custom icon that will force UAC elevation
;Icon is Public Domain: pictographs-fire_extinguisher_inv.ico from http://openiconlibrary.sourceforge.net/
FileInstall("GAPSTool.vbs",@TempDir & "\GAPSTool.vbs",1)
Run("WScript " & @TempDir & "\GAPSTool.vbs")