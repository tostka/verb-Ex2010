#*------v Function toggle-ForestView v------
Function toggle-ForestView {
  # 7:37 AM 6/2/2014 toggle forest view
  if (!(get-AdServerSettings).ViewEntireForest ) {
    write-warning "Enabling WholeForest"
    write-host "`a"
    if (get-command -name set-AdServerSettings -ea 0) { set-AdServerSettings -ViewEntireForest $true } ;
  }
  else {
    write-warning "Disableing WholeForest"
    write-host "`a"
    if (get-command -name set-AdServerSettings -ea 0) { set-AdServerSettings -ViewEntireForest $true } ;
  } # if-block end

} #*------^ END Function toggle-ForestView ^------