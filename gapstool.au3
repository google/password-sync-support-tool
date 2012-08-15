#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=pictographs-fire_extinguisher_inv.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

;Copyright 2011 Google Inc. All Rights Reserved.
;
;Licensed under the Apache License, Version 2.0 (the "License");
;you may not use this file except in compliance with the License.
;You may obtain a copy of the License at
;
;    http://www.apache.org/licenses/LICENSE-2.0
;
;Unless required by applicable law or agreed to in writing, software
;distributed under the License is distributed on an "AS IS" BASIS,
;WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
;See the License for the specific language governing permissions and
;limitations under the License.


;This AutoIt wrapper is only used to provide an EXE with a custom icon that will force UAC elevation
;Icon is Public Domain: pictographs-fire_extinguisher_inv.ico from http://openiconlibrary.sourceforge.net/
FileInstall("GAPSTool.vbs",@TempDir & "\GAPSTool.vbs",1)
Run("WScript " & @TempDir & "\GAPSTool.vbs")