<##
Surveillance multizones MORPHEE (.xcce) – v2.9.1 (PS 5.1)

Nouveaux points vs 2.9.0:
- Détection zones explicites plus robuste:
  * TRIM systématique de ZONES_SPEC
  * STRICT_MATCH optionnel (balises ancrées ^ $)
  * Échappement sûr pour git -L (incluant '/')
  * Fallback numérique si git -L par balises ne renvoie rien (balises renommées)
  * Warnings sur START/END orphelins ou déséquilibrés
  * zoneid univoque si plusieurs définitions partagent le même Name
- Conserve 2.9.0:
  * CSV “courant” réinitialisé
  * Colonne "modifié le" = date/heure de la zone (commit/local)
  * UI: explicit affiché en CERTIF, accents sûrs
  * Rapport HTML groupé par fichier (repliable), filtres, export CSV filtré
##>

# Paramètres configurables (valeurs par défaut = comportement actuel)
param(
  [string]$FOLDER           = 'D:\Bitbucket\test_noe_global_standard_ref',
  [string]$EXTENSIONS       = '*.xcce',                # ex: "*.xcce;*.xml"
  [int]   $RETENTION_DAYS   = 30,
  [string]$ZONES_SPEC       = 'GEN|### START_ZONE_A_SURVEILLER|### END_ZONE_A_SURVEILLER',  # "Nom|Start|End;..."
  [ValidateSet('HEAD','START')]
  [string]$LOCAL_BASE       = 'START',                 # 'HEAD' ou 'START'
  [bool]  $DEBUG_ZONES      = $false,

  # Hash / identifiants
  [int]   $PATHHASH_HEXLEN  = 12,
  [ValidateSet('ON','OFF')]
  [string]$CONTENTHASH_MODE = 'ON',

  # Fallback zones auto (<Step>/AUTO_FILE) ?
  [ValidateSet('OFF','ON')]
  [string]$AUTO_ZONES       = 'OFF',                   # 'OFF' = uniquement zones explicites. 'ON' pour activer les zones auto.

  # Correspondance stricte des balises ?
  # OFF = "contient le texte" (souple, par défaut) ; ON = ligne entière (ancrage ^ $)
  [ValidateSet('OFF','ON')]
  [string]$STRICT_MATCH     = 'OFF'                    # 'OFF' ou 'ON'
)

# =======================
# 1) Configuration
# =======================
# Les variables ci-dessus sont désormais fournies via le bloc param
$FOLDER           = $FOLDER
$EXTENSIONS       = $EXTENSIONS                 # ex: "*.xcce;*.xml"
$RETENTION_DAYS   = $RETENTION_DAYS
$ZONES_SPEC       = $ZONES_SPEC                  # "Nom|Start|End;..."
$LOCAL_BASE       = $LOCAL_BASE                  # 'HEAD' ou 'START'
$DEBUG_ZONES      = $DEBUG_ZONES

# Hash / identifiants
$PATHHASH_HEXLEN  = $PATHHASH_HEXLEN
$CONTENTHASH_MODE = $CONTENTHASH_MODE            # 'ON'/'OFF'

# Fallback zones auto (<Step>/AUTO_FILE) ?
$AUTO_ZONES       = $AUTO_ZONES                  # 'OFF' = uniquement zones explicites. 'ON' pour activer les zones auto.

# Correspondance stricte des balises ?
# OFF = "contient le texte" (souple, par défaut) ; ON = ligne entière (ancrage ^ $)
$STRICT_MATCH     = $STRICT_MATCH                # 'OFF' ou 'ON'

# =======================
# 2) Consoles/Encodages
# =======================
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch {}

# =======================
# 3) Utilitaires
# =======================
function Write-Info($m){ Write-Host "[INFO] $m" -ForegroundColor Cyan }
function Write-Warn($m){ Write-Host "[WARN] $m" -ForegroundColor Yellow }
function Write-Err ($m){ Write-Host "[ERROR] $m" -ForegroundColor Red }

function Get-RunStamp(){ (Get-Date).ToString('yyyyMMdd_HHmmss') }
function Get-DateFolder(){ (Get-Date).ToString('yyyy_MM_dd') }
function Get-DateYmd(){ (Get-Date).ToString('yyyyMMdd') }

function Ensure-Tool($name){
  try { Get-Command $name -ErrorAction Stop | Out-Null; $true }
  catch { throw "Outil manquant: $name" }
}

function Test-GitRepo($path){
  & git -C $path rev-parse --is-inside-work-tree *> $null
  return ($LASTEXITCODE -eq 0)
}

function Get-PathHashN([string]$relPath,[int]$hlen){
  $bytes = [Text.Encoding]::UTF8.GetBytes($relPath)
  $md5   = [Security.Cryptography.MD5]::Create().ComputeHash($bytes)
  $hex   = -join ($md5 | ForEach-Object { $_.ToString('x2') })
  return $hex.Substring(0,[Math]::Min($hlen,$hex.Length))
}

function Get-ContentHashN([string]$full,[int]$hlen){
  if(-not (Test-Path $full)){ return '-' }
  try{
    $h = (Get-FileHash -LiteralPath $full -Algorithm MD5).Hash.ToLower()
    return $h.Substring(0,[Math]::Min($hlen,$h.Length))
  } catch {
    return '-'
  }
}

function Join-PathSafe([string]$root, [string]$rel){
  $rel = $rel -replace '/', '\\'
  return (Join-Path $root $rel)
}

# -> "2025-08-12 00:49:15"
function To-HumanLocal([string]$iso){
  try{
    $dto = [datetimeoffset]::Parse($iso, [System.Globalization.CultureInfo]::InvariantCulture)
    return $dto.ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')
  } catch { return $iso }
}

# Echappe une balise pour git -L "/.../,/.../"
# - [regex]::Escape ne gère pas le '/' -> on l'échappe aussi
function Escape-GitL([string]$s){
  if([string]::IsNullOrEmpty($s)){ return $s }
  return ([regex]::Escape($s)).Replace('/','\/')
}

# Dernière date de commit qui a touché la zone (avec fallback numérique et fichier)
function Get-ZoneLastCommitIso([string]$relPath,[int]$zstart,[int]$zend,[string]$zstartlab,[string]$zendlab){
  $iso = $null
  if($zstartlab -and $zstartlab -ne 'AUTO_NUMERIC'){
    $sl = Escape-GitL $zstartlab
    $el = Escape-GitL $zendlab
    $iso = (& git -c core.quotepath=false -C $FOLDER log "$START_COMMIT..HEAD" --date=iso-strict --pretty=format:"%cI" -L "/$sl/,/$el/:$relPath" 2>$null | Select-Object -First 1)
  }
  if([string]::IsNullOrWhiteSpace($iso)){
    $iso = (& git -c core.quotepath=false -C $FOLDER log "$START_COMMIT..HEAD" --date=iso-strict --pretty=format:"%cI" -L ("{0},{1}:{2}" -f $zstart,$zend,$relPath) 2>$null | Select-Object -First 1)
  }
  if([string]::IsNullOrWhiteSpace($iso)){
    $iso = (& git -c core.quotepath=false -C $FOLDER log -1 --date=iso-strict --pretty=format:"%cI" -- $relPath 2>$null)
  }
  return $iso
}

# =======================
# 4) Préparation & validations
# =======================
$RUN_STAMP   = Get-RunStamp
$DATE_FOLDER = Get-DateFolder
$DATE_YMD    = $RUN_STAMP.Substring(0,8)

[void](Ensure-Tool git)
Write-Info "Depot            : '$FOLDER'"
Write-Info "Extensions       : $EXTENSIONS"
Write-Info "Zones (explicites): $ZONES_SPEC"
Write-Info "Local base       : $LOCAL_BASE  (HEAD ou START)"
Write-Info "Fallback auto     : $AUTO_ZONES (OFF = uniquement explicites)"
Write-Info "Strict match      : $STRICT_MATCH (ON = balises ancrées ^ $)"
Write-Info "Retention logs   : $RETENTION_DAYS jours"
$startTime = Get-Date
Write-Info "Debut : $($startTime.ToString('HH:mm:ss.fff')) (stamp $RUN_STAMP)"

if(-not (Test-GitRepo $FOLDER)){
  Write-Err "'$FOLDER' n'est pas un depot Git"; exit 1
}

Write-Info 'git fetch --tags...'
& git -C $FOLDER fetch --tags *> $null
if($LASTEXITCODE -ne 0){ Write-Warn 'fetch --tags a echoue - hors-ligne ou pas de remote - on continue' } else { Write-Info 'fetch --tags OK' }

# START_COMMIT depuis tag Certif-global-V* (le plus récent), sinon root commit
Write-Info "Recherche tag 'Certif-global-V*'..."
$globalTag = (& git -C $FOLDER tag --list 'Certif-global-V*' --sort=-v:refname | Select-Object -First 1)
$START_COMMIT = $null
if([string]::IsNullOrWhiteSpace($globalTag)){
  Write-Warn 'Aucun tag global -> premier commit'
  $START_COMMIT = (& git -C $FOLDER rev-list --max-parents=0 HEAD).Trim()
}else{
  Write-Info "Tag global: $globalTag"
  $START_COMMIT = (& git -C $FOLDER rev-list -n1 $globalTag).Trim()
}
Write-Info "START_COMMIT: $START_COMMIT"

# =======================
# 5) Logs + purge
# =======================
$LOG_ROOT = Join-Path $FOLDER 'logs'
$LOG_DIR  = Join-Path $LOG_ROOT $DATE_FOLDER
if(-not (Test-Path $LOG_DIR)){ New-Item -ItemType Directory -Path $LOG_DIR -Force | Out-Null }

# Purge dossiers logs par rétention
Get-ChildItem -Path $LOG_ROOT -Directory -ErrorAction SilentlyContinue | ForEach-Object {
  $dname = $_.Name
  $culture = [System.Globalization.CultureInfo]::InvariantCulture
  [datetime]$dt = [datetime]::MinValue
  if([datetime]::TryParseExact($dname,'yyyy_MM_dd',$culture,[System.Globalization.DateTimeStyles]::None,[ref]$dt)){
    if((Get-Date) - $dt -gt [TimeSpan]::FromDays($RETENTION_DAYS)){
      try{ Remove-Item -Recurse -Force -Path $_.FullName }catch{}
    }
  }
}

# =======================
# 6) CSV "courant" (réinitialisé à chaque run)
# =======================
$CSV_FILE = Join-Path $LOG_DIR 'changes_current.csv'
$EXPECTED_HDR = 'run_ts;zone_ts;type;zoneid;zoneclass;filepath;pathhash;contenthash;logfile;start;end'
# on écrase systématiquement
$EXPECTED_HDR | Out-File -FilePath $CSV_FILE -Encoding UTF8 -Force

# Compteurs
$script:COUNT_TOTAL=0
$script:COUNT_CHANGED=0
$script:COUNT_DELETED=0
$script:COUNT_ERRORS=0
$script:COUNT_CHANGED_EXPLICIT=0
$script:COUNT_CHANGED_AUTO=0
$script:COUNT_IGNORED_NO_EXPLICIT=0
$script:ERROR_MESSAGES = [System.Text.StringBuilder]::new()

# =======================
# 7) Collecte des candidats
# =======================
Write-Info 'Prefiltrage candidats...'
$candsRaw = @()
$candsRaw += (& git -c core.quotepath=false -C $FOLDER diff "$START_COMMIT..HEAD" --diff-filter=AM --name-only) 2>$null
$candsRaw += (& git -c core.quotepath=false -C $FOLDER diff "$START_COMMIT..HEAD" --diff-filter=D  --name-only) 2>$null
$candsRaw += (& git -c core.quotepath=false -C $FOLDER diff                 --diff-filter=AM --name-only) 2>$null     # unstaged
$candsRaw += (& git -c core.quotepath=false -C $FOLDER diff --cached        --diff-filter=AM --name-only) 2>$null     # staged
$candsRaw += (& git -c core.quotepath=false -C $FOLDER diff                 --diff-filter=D  --name-only) 2>$null
$candsRaw += (& git -c core.quotepath=false -C $FOLDER diff --cached        --diff-filter=D  --name-only) 2>$null
$candsRaw += (& git -c core.quotepath=false -C $FOLDER ls-files --others --exclude-standard) 2>$null

# Filtrage extensions
$patterns = $EXTENSIONS.Split(';',[StringSplitOptions]::RemoveEmptyEntries) |
  ForEach-Object { New-Object System.Management.Automation.WildcardPattern($_,[System.Management.Automation.WildcardOptions]::IgnoreCase) }

$candidates = $candsRaw | Where-Object {
  $p = $_; if([string]::IsNullOrWhiteSpace($p)){ return $false }
  foreach($wp in $patterns){ if($wp.IsMatch($p)){ return $true } }
  $false
} | Sort-Object -Unique

if($candidates.Count -eq 0){
  Write-Info 'Aucun candidat via diff -> fallback liste complete...'
  $allList = @()
  $allList += (& git -c core.quotepath=false -C $FOLDER ls-files) 2>$null
  $allList += (& git -c core.quotepath=false -C $FOLDER ls-files --others --exclude-standard) 2>$null
  $candidates = $allList | Where-Object {
    $p = $_; if([string]::IsNullOrWhiteSpace($p)){ return $false }
    foreach($wp in $patterns){ if($wp.IsMatch($p)){ return $true } }
    $false
  } | Sort-Object -Unique
}
Write-Info ("Fichiers candidats: {0}" -f $candidates.Count)
$candidates | Select-Object -First 10 | ForEach-Object { Write-Host "[CAND] $_" }

# =======================
# 8) Fonctions coeur
# =======================
function Enumerate-Zones([string]$relPath){
  $full = Join-PathSafe $FOLDER $relPath
  if(-not (Test-Path $full)){ return @() }
  try { $txt = Get-Content -LiteralPath $full -Encoding utf8 } catch { $txt = Get-Content -LiteralPath $full }
  $rows = @()

  # Zones explicites (parsing robuste + TRIM)
  $defs = @()
  foreach($item in $ZONES_SPEC.Split(';',[System.StringSplitOptions]::RemoveEmptyEntries)){
    if($item -match '\|'){
      $parts = $item.Split('|',3) | ForEach-Object { $_.Trim() }
      $defs += [pscustomobject]@{ Name=$parts[0]; S=$parts[1]; E=$parts[2] }
    }
  }

  # Compter occurrences par Name (pour zoneid univoques si besoin)
  $nameCounts = @{}
  foreach($d in $defs){
    $key = $d.Name.ToLowerInvariant()
    if(-not $nameCounts.ContainsKey($key)){ $nameCounts[$key]=0 }
    $nameCounts[$key] = $nameCounts[$key] + 1
  }

  $defIdx = 0
  foreach($d in $defs){
    $defIdx++

    # Sélection des lignes Start/End: strict (ancré) ou souple
    if($STRICT_MATCH -eq 'ON'){
      $reS = '^\s*' + [regex]::Escape($d.S) + '\s*$'
      $reE = '^\s*' + [regex]::Escape($d.E) + '\s*$'
      $starts = Select-String -Path $full -Pattern $reS | ForEach-Object LineNumber | Sort-Object
      $ends   = Select-String -Path $full -Pattern $reE | ForEach-Object LineNumber | Sort-Object
    } else {
      $starts = Select-String -Path $full -SimpleMatch $d.S -AllMatches | ForEach-Object LineNumber | Sort-Object
      $ends   = Select-String -Path $full -SimpleMatch $d.E -AllMatches | ForEach-Object LineNumber | Sort-Object
    }

    if($starts.Count -ne $ends.Count){
      Write-Warn ("Balises déséquilibrées pour '{0}' dans '{1}' (starts={2}, ends={3})" -f $d.Name,$relPath,$starts.Count,$ends.Count)
    }

    $i=0; $j=0; $idx=0
    while($i -lt $starts.Count){
      $s=$starts[$i]
      while($j -lt $ends.Count -and $ends[$j] -le $s){ $j++ }
      if($j -lt $ends.Count){
        $e=$ends[$j]; $idx++

        $needDisamb = $nameCounts[$d.Name.ToLowerInvariant()] -gt 1
        $zoneIdCore = if($needDisamb){ ('{0}d{1}_{2}' -f $d.Name,$defIdx,$idx) } else { ('{0}_{1}' -f $d.Name,$idx) }

        $rows += [pscustomobject]@{
          zoneid     = $zoneIdCore
          start      = $s
          end        = $e
          name       = $d.Name
          index      = $idx
          startLabel = $d.S
          endLabel   = $d.E
          zoneClass  = 'explicit'
        }
        $j++; $i++
      } else {
        Write-Warn ("START sans END pour '{0}' à partir de la ligne {1} dans '{2}'" -f $d.Name,$s,$relPath)
        break
      }
    }
  }

  if($rows.Count -gt 0){ return $rows }
  if($AUTO_ZONES -ne 'ON'){ return @() }

  # Fallback <Step>
  $openLines  = Select-String -Path $full -Pattern '^\s*<Step(\s|>)'   | ForEach-Object LineNumber
  $closeLines = Select-String -Path $full -Pattern '^\s*</Step>'        | ForEach-Object LineNumber
  $n = [Math]::Min($openLines.Count,$closeLines.Count)
  $idx=0
  for($k=0; $k -lt $n; $k++){
    $s=$openLines[$k]; $e=$closeLines[$k]
    if($e -ge $s){
      $idx++
      $block = ($txt[($s-1)..($e-1)] -join [Environment]::NewLine)
      $label = 'Step_' + $idx
      $m = [regex]::Match($block,'<Label>\s*(.*?)\s*</Label>',[Text.RegularExpressions.RegexOptions]::Singleline)
      if($m.Success){ $label = $m.Groups[1].Value.Trim() }
      $rows += [pscustomobject]@{
        zoneid=("AUTO_STEP_{0}" -f $idx); start=$s; end=$e; name=$label; index=$idx;
        startLabel='AUTO_NUMERIC'; endLabel='AUTO_NUMERIC'; zoneClass='auto_step'
      }
    }
  }
  if($rows.Count -eq 0){
    $lineCount = $txt.Count; if($lineCount -lt 1){ $lineCount = 1 }
    $rows += [pscustomobject]@{
      zoneid='AUTO_FILE_1'; start=1; end=$lineCount; name='AUTO_FILE'; index=1;
      startLabel='AUTO_NUMERIC'; endLabel='AUTO_NUMERIC'; zoneClass='auto_file'
    }
  }
  return $rows
}

function StrictLocal-Detect([string]$relPath, $zones){
  $full = Join-PathSafe $FOLDER $relPath
  $hits = New-Object System.Collections.Generic.HashSet[string]

  & git -c core.quotepath=false -C $FOLDER ls-files --error-unmatch -- $relPath *> $null
  $tracked = ($LASTEXITCODE -eq 0)

  if(-not (Test-Path $full)){ return @() }

  if(-not $tracked){
    $N = (Get-Content -LiteralPath $full).Count
    if($N -lt 1){ $N = 1 }
    foreach($z in $zones){
      if(1 -le $z.end -and $N -ge $z.start){ [void]$hits.Add($z.zoneid) }
    }
    return $hits
  }

  $base = if($LOCAL_BASE -eq 'START' -and $START_COMMIT){ $START_COMMIT } else { 'HEAD' }
  $diff = (& git -c core.quotepath=false -C $FOLDER diff -U0 $base -- $relPath 2>$null)
  $hunks = $diff | Select-String '^@@\s-\d+(?:,\d+)?\s\+\d+(?:,\d+)?\s@@'

  foreach($h in $hunks){
    if($h.Line -match '^@@\s-(?<os>\d+)(,(?<ol>\d+))?\s\+(?<ns>\d+)(,(?<nl>\d+))?\s@@'){
      $os=[int]$Matches.os; $ol=[int]$Matches.ol; if($ol -le 0){$ol=1}; $oe=$os+$ol-1
      $ns=[int]$Matches.ns; $nl=[int]$Matches.nl; if($nl -le 0){$nl=1}; $ne=$ns+$nl-1
      foreach($z in $zones){
        $overOld = ($os -le $z.end -and $oe -ge $z.start)
        $overNew = ($ns -le $z.end -and $ne -ge $z.start)
        if($overOld -or $overNew){ [void]$hits.Add($z.zoneid) }
      }
    }
  }
  return $hits
}

function Log-Change([string]$relPath,[string]$type,[string]$zoneid,[string]$zoneclass,[int]$zstart,[int]$zend,[string]$zstartlab,[string]$zendlab){
  $relWin    = $relPath -replace '/', '\\'
  $basename  = [IO.Path]::GetFileNameWithoutExtension($relWin)
  $pathhash  = Get-PathHashN $relWin $PATHHASH_HEXLEN
  $full      = Join-PathSafe $FOLDER $relPath
  $stamp     = Get-RunStamp

  $contenthash = '-'
  if($CONTENTHASH_MODE -eq 'ON' -and $type -ne 'deleted'){
    $contenthash = Get-ContentHashN $full $PATHHASH_HEXLEN
  }

  # calcul "zone_ts"
  $zone_ts = ''
  if($type -ieq 'commit'){
    $iso = Get-ZoneLastCommitIso $relPath $zstart $zend $zstartlab $zendlab
    if([string]::IsNullOrWhiteSpace($iso)){ $zone_ts = $stamp -replace '_',' ' } else { $zone_ts = To-HumanLocal $iso }
  } elseif($type -ieq 'local'){
    try{ $zone_ts = (Get-Item -LiteralPath $full).LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss') } catch { $zone_ts = $stamp -replace '_',' ' }
  } elseif($type -ieq 'deleted'){
    $iso = (& git -c core.quotepath=false -C $FOLDER log -1 --date=iso-strict --pretty=format:"%cI" -- $relPath 2>$null)
    if([string]::IsNullOrWhiteSpace($iso)){ $zone_ts = $stamp -replace '_',' ' } else { $zone_ts = To-HumanLocal $iso }
  }

  $invalidChars = [IO.Path]::GetInvalidFileNameChars() + [IO.Path]::DirectorySeparatorChar + [IO.Path]::AltDirectorySeparatorChar
  $invalidStr   = -join $invalidChars
  $re           = '[' + ([Regex]::Escape($invalidStr)) + ']'
  $zoneSafe  = ($zoneid  -replace $re,'_')
  $baseSafe  = ($basename -replace $re,'_')
  $classSafe = ($zoneclass -replace $re,'_')

  $logFile  = Join-Path $LOG_DIR ("log_output_{0}_{1}_{2}_{3}_{4}_{5}.txt" -f $baseSafe,$pathhash,$classSafe,$zoneSafe,$type,$stamp)

  if($type -ieq 'deleted'){ $script:COUNT_DELETED++ } else { $script:COUNT_CHANGED++ }
  if($zoneclass -eq 'explicit'){ $script:COUNT_CHANGED_EXPLICIT++ } else { $script:COUNT_CHANGED_AUTO++ }
  Write-Info ([string]::Format('[{0}][{1}] {2} -> "{3}"  (zone_ts={4})',$type,$zoneclass,$zoneid,$relPath,$zone_ts))

  try{
    if($type -ieq 'commit'){
      if($zstartlab -and $zstartlab -ne 'AUTO_NUMERIC'){
        $sl = Escape-GitL $zstartlab; $el = Escape-GitL $zendlab
        $out = (& git -c core.quotepath=false -C $FOLDER log "$START_COMMIT..HEAD" -p -L "/$sl/,/$el/:$relPath" 2>$null)
        if(-not $out){
          $out = (& git -c core.quotepath=false -C $FOLDER log "$START_COMMIT..HEAD" -p -L ("{0},{1}:{2}" -f $zstart,$zend,$relPath) 2>$null)
        }
        $out | Out-File -FilePath $logFile -Encoding UTF8 -Force
      } else {
        & git -c core.quotepath=false -C $FOLDER log "$START_COMMIT..HEAD" -p -L ("{0},{1}:{2}" -f $zstart,$zend,$relPath) 2>$null | Out-File -FilePath $logFile -Encoding UTF8 -Force
      }
    } elseif($type -ieq 'local'){
      if($LOCAL_BASE -ieq 'START'){
        & git -c core.quotepath=false -C $FOLDER diff -U0 $START_COMMIT -- $relPath 2>$null | Out-File -FilePath $logFile -Encoding UTF8 -Force
      } else {
        & git -c core.quotepath=false -C $FOLDER diff -U0 HEAD -- $relPath 2>$null | Out-File -FilePath $logFile -Encoding UTF8 -Force
      }
    } elseif($type -ieq 'deleted'){
      & git -c core.quotepath=false -C $FOLDER diff "$START_COMMIT..HEAD" -- $relPath 2>$null | Out-File -FilePath $logFile -Encoding UTF8 -Force
    }
  } catch {
    Write-Err ("echec generation log {0}" -f $basename)
    $script:COUNT_ERRORS++
    [void]$script:ERROR_MESSAGES.Append('log ').Append($basename).Append(';')
  }

  $run_ts = Get-RunStamp
  $csvLine = '{0};{1};{2};{3};{4};"{5}";{6};{7};"{8}";{9};{10}' -f $run_ts,$zone_ts,$type,$zoneid,$zoneclass,$relPath,$pathhash,$contenthash,$logFile,$zstart,$zend
  Add-Content -Path $CSV_FILE -Value $csvLine -Encoding UTF8
}

# =======================
# 9) Analyse par fichier
# =======================
$prevFile = $null
foreach($f in $candidates){
  if($f -ne $prevFile){
    $prevFile = $f
    $script:COUNT_TOTAL++
    Write-Info ("---- '{0}'" -f $f)

    $full = Join-PathSafe $FOLDER $f
    if(-not (Test-Path $full)){
      Log-Change $f 'deleted' 'NA' 'n/a' 0 0 '' ''
      continue
    }

    $zones = Enumerate-Zones $f
    Write-Info ("Zones trouvees : {0}" -f (@($zones).Count))

    if($zones.Count -eq 0){
      if($AUTO_ZONES -ne 'ON'){
        Write-Info "Aucune zone explicite -> fichier ignore (AUTO_ZONES=$AUTO_ZONES)"
        $script:COUNT_IGNORED_NO_EXPLICIT++
        continue
      } else {
        Write-Warn 'Aucune zone explicite — fallback AUTO_* appliqué (AUTO_ZONES=ON)'
      }
    }

    # fichier touché historiquement ?
    $fileTouchedHist = $false
    $probeHist = (& git -c core.quotepath=false -C $FOLDER diff "$START_COMMIT..HEAD" --name-only -- $f 2>$null)
    if($probeHist){ $fileTouchedHist = $true }

    if($fileTouchedHist){
      foreach($z in $zones){
        $hasCommit = $false
        if($z.startLabel -ne 'AUTO_NUMERIC'){
          $sl = Escape-GitL $z.startLabel
          $el = Escape-GitL $z.endLabel
          $probe = (& git -c core.quotepath=false -C $FOLDER log "$START_COMMIT..HEAD" -p -L "/$sl/,/$el/:$f" 2>$null)
          if(-not $probe){
            $probe = (& git -c core.quotepath=false -C $FOLDER log "$START_COMMIT..HEAD" -p -L ("{0},{1}:{2}" -f $z.start,$z.end,$f) 2>$null)
          }
          if($probe | Select-String '^commit ' -SimpleMatch){ $hasCommit = $true }
        } else {
          $probe = (& git -c core.quotepath=false -C $FOLDER log "$START_COMMIT..HEAD" -p -L ("{0},{1}:{2}" -f $z.start,$z.end,$f) 2>$null)
          if($probe | Select-String '^commit ' -SimpleMatch){ $hasCommit = $true }
        }
        if($hasCommit){
          if($z.startLabel -ne 'AUTO_NUMERIC'){
            Log-Change $f 'commit' $z.zoneid $z.zoneClass $z.start $z.end $z.startLabel $z.endLabel
          } else {
            Log-Change $f 'commit' $z.zoneid $z.zoneClass $z.start $z.end '' ''
          }
        }
      }
    }

    # Détection locale stricte
    $localHits = StrictLocal-Detect $f $zones
    foreach($lz in $localHits){
      $z = $zones | Where-Object zoneid -eq $lz | Select-Object -First 1
      if($null -ne $z){ Log-Change $f 'local' $z.zoneid $z.zoneClass $z.start $z.end '' '' }
    }
  }
}

# =======================
# 10) Fin + Summary + Rapport HTML
# =======================
$endTime = Get-Date
Write-Info ("Fin : {0}" -f $endTime.ToString('HH:mm:ss.fff'))
Write-Info ("Fichiers examines             : {0}" -f $script:COUNT_TOTAL)
Write-Info ("Modifies (tous)               : {0}" -f $script:COUNT_CHANGED)
Write-Info ("  dont explicit               : {0}" -f $script:COUNT_CHANGED_EXPLICIT)
Write-Info ("  dont auto                   : {0}" -f $script:COUNT_CHANGED_AUTO)
Write-Info ("Supprimes                     : {0}" -f $script:COUNT_DELETED)
Write-Info ("Ignorés (sans zone explicite) : {0}" -f $script:COUNT_IGNORED_NO_EXPLICIT)
Write-Info ("Erreurs                       : {0}" -f $script:COUNT_ERRORS)
Write-Info ("Duree                         : {0} -> {1}" -f $startTime.ToString('HH:mm:ss.fff'), $endTime.ToString('HH:mm:ss.fff'))

$summary = @(
  "Resume du $DATE_FOLDER de $($startTime.ToString('HH:mm:ss.fff')) a $($endTime.ToString('HH:mm:ss.fff'))",
  "Fichiers examines             : $script:COUNT_TOTAL",
  "Modifies (tous)               : $script:COUNT_CHANGED  | explicit: $script:COUNT_CHANGED_EXPLICIT | auto: $script:COUNT_CHANGED_AUTO",
  "Supprimes                     : $script:COUNT_DELETED",
  "Ignorés (sans zone explicite) : $script:COUNT_IGNORED_NO_EXPLICIT",
  "Erreurs                       : $script:COUNT_ERRORS",
  ("Messages d'erreur : {0}" -f $script:ERROR_MESSAGES.ToString())
)
$summaryPath = Join-Path $LOG_DIR ("summary_{0}.txt" -f $RUN_STAMP)
$summary | Out-File -FilePath $summaryPath -Encoding UTF8 -Force

# -------- HTML (utilise zone_ts, CSV courant) --------
function Build-HtmlReport(){
  param([string]$Csv,[string]$OutDir,[string]$DateYmd)

  $rows = Import-Csv -Delimiter ';' -LiteralPath $Csv
  # groupement par fichier
  $groups = $rows | Group-Object filepath | Sort-Object Name

  # y a-t-il des zones CERTIF dans ce run ?
  $hasCertif = $rows | Where-Object { $_.zoneclass -eq 'explicit' } | Select-Object -First 1
  $selExpJs = if($null -ne $hasCertif){ 'true' } else { 'false' }

  $header = @'
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Rapport changements</title>
<style>
:root{
  --bg:#ffffff; --fg:#1f2937; --muted:#6b7280; --line:#e5e7eb; --th-bg:#f3f4f6; --accent:#2563eb; --badge:#111827; --badge-bg:#e5e7eb;
  --commit:#065f46; --local:#7c2d12; --deleted:#991b1b; --chip-bg:#e5e7eb;
}
.dark{
  --bg:#0b1020; --fg:#e5e7eb; --muted:#9ca3af; --line:#1f2937; --th-bg:#111827; --accent:#60a5fa; --badge:#e5e7eb; --badge-bg:#374151;
  --commit:#34d399; --local:#f59e0b; --deleted:#f87171; --chip-bg:#374151;
}
*{box-sizing:border-box}
body{font-family:Segoe UI,Arial,sans-serif;background:var(--bg);color:var(--fg);margin:0}
.header{position:sticky;top:0;background:var(--bg);z-index:5;border-bottom:1px solid var(--line)}
.header-inner{display:flex;align-items:center;gap:12px; padding:10px 12px}
.h-title{font-size:18px;font-weight:600}

/* barre outils */
.toolbar{position:sticky;top:48px;background:var(--bg);z-index:4;border-bottom:1px solid var(--line);padding:8px 12px;display:flex;flex-wrap:wrap;gap:8px;align-items:center}
.toolbar label{display:flex;align-items:center;gap:6px}
input[type="text"], select{padding:6px 8px;border:1px solid var(--line);background:transparent;color:var(--fg)}
input[type="checkbox"]{transform:translateY(1px)}
button{padding:6px 10px;border:1px solid var(--line);background:var(--chip-bg);color:var(--fg);cursor:pointer}
button:hover{border-color:var(--accent)}
.badge{display:inline-block;padding:2px 8px;border-radius:999px;background:var(--badge-bg);color:var(--badge);font-size:12px;margin-left:6px}

/* groupes par fichier */
.group{border-bottom:1px solid var(--line)}
.group>summary{list-style:none; cursor:pointer; padding:10px 12px; display:flex; align-items:center; gap:10px; background:var(--bg); position:sticky; top:96px; z-index:3; border-bottom:1px dashed var(--line)}
.group[open]>summary{border-bottom-color:transparent}
.g-title{font-weight:600}
.g-badges .badge{margin-left:4px}

/* tables groupe */
table{border-collapse:collapse;width:100%}
th,td{border:1px solid var(--line);padding:6px 8px;vertical-align:top}
thead th{background:var(--th-bg);cursor:pointer}
tbody tr:nth-child(even){background-color:rgba(0,0,0,0.02)}
a{color:var(--accent);text-decoration:none}
a:hover{text-decoration:underline}
.type-commit td{border-left:4px solid var(--commit)}
.type-local td{border-left:4px solid var(--local)}
.type-deleted td{border-left:4px solid var(--deleted)}

.help{padding:10px 12px;border-bottom:1px solid var(--line);background:var(--bg)}
.help summary{cursor:pointer;font-weight:600}
@media (max-width:1000px){ .opt-cols{display:none} }
</style>
<script>
(function(){
  var STATE_KEY='morphee_report_filters_v291';
  var DATEYMD='__DATEYMD__';
  var AUTO_ZONES='__AUTO_ZONES__';
  var PREF_CERTIF=__PREF_CERTIF__;

  function $(id){return document.getElementById(id)}
  function $all(sel,root){return (root||document).querySelectorAll(sel)}
  function debounce(fn,ms){var t;return function(){var a=arguments,c=this;clearTimeout(t);t=setTimeout(function(){fn.apply(c,a)},ms)}}

  function saveState(){
    var st={
      fz:$('fz').value, ff:$('ff').value, ft:$('ft').value, fc:$('fc').value,
      todayOnly:$('todayOnly').checked, regex:$('regex').checked, dark:document.body.classList.contains('dark'),
      colPath:$('colPath').checked, colHash:$('colHash').checked, colCHash:$('colCHash').checked, colStart:$('colStart').checked, colEnd:$('colEnd').checked
    };
    try{localStorage.setItem(STATE_KEY,JSON.stringify(st))}catch(e){}
  }
  function loadState(){
    try{
      var raw=localStorage.getItem(STATE_KEY); if(!raw) return;
      var st=JSON.parse(raw);
      if(st.fz!==undefined)$('fz').value=st.fz;
      if(st.ff!==undefined)$('ff').value=st.ff;
      if(st.ft!==undefined)$('ft').value=st.ft;
      if(st.fc!==undefined)$('fc').value=st.fc;
      if(st.todayOnly!==undefined)$('todayOnly').checked=st.todayOnly;
      if(st.regex!==undefined)$('regex').checked=st.regex;
      if(st.dark){document.body.classList.add('dark');$('dark').checked=true}
      if(st.colPath!==undefined)$('colPath').checked=st.colPath;
      if(st.colHash!==undefined)$('colHash').checked=st.colHash;
      if(st.colCHash!==undefined)$('colCHash').checked=st.colCHash;
      if(st.colStart!==undefined)$('colStart').checked=st.colStart;
      if(st.colEnd!==undefined)$('colEnd').checked=st.colEnd;
    }catch(e){}
  }

  function applyColVisibility(){
    var vis=[true,true,true,true,$('colPath').checked,$('colHash').checked,$('colCHash').checked,true,$('colStart').checked,$('colEnd').checked,true];
    $all('.group-table').forEach(function(tbl){
      var h=tbl.tHead.rows[0].cells;
      for(var j=0;j<vis.length;j++){ if(h[j]) h[j].style.display=vis[j]?'':'none'; }
      var r=tbl.tBodies[0].rows;
      for(var i=0;i<r.length;i++){ for(var j=0;j<vis.length;j++){ var c=r[i].cells[j]; if(c) c.style.display=vis[j]?'':'none'; } }
    });
  }

  // tri pour un *groupe* (table spécifique)
  function sg(tableId,n){
    var t=$(tableId); if(!t) return;
    var r=t.tBodies[0].rows, sw=1;
    while(sw){
      sw=0;
      for(var i=0;i<r.length-1;i++){
        if(r[i].style.display==='none') continue;
        var j=i+1; while(j<r.length && r[j].style.display==='none') j++;
        if(j>=r.length) break;
        var x=r[i].cells[n].innerText, y=r[j].cells[n].innerText;
        var xn=parseFloat(x), yn=parseFloat(y); var cmp;
        if(!isNaN(xn)&&!isNaN(yn)){ cmp = xn>yn; } else { cmp = x.toLowerCase()>y.toLowerCase(); }
        if(cmp){ r[i].parentNode.insertBefore(r[j],r[i]); sw=1; }
      }
    }
  }
  window.sg=sg;

  function filterRows(){
    var vz=$('fz').value.toLowerCase().trim();
    var vf=$('ff').value.toLowerCase().trim();
    var ty=$('ft').value.toLowerCase();
    var cl=$('fc').value.toLowerCase(); // 'certif', 'auto_step', 'auto_file' ou ''
    var todayOnly=$('todayOnly').checked;
    var dy=DATEYMD;
    var useRegex=$('regex').checked;

    var reZ=null, reF=null;
    if(useRegex){
      try{ if(vz) reZ=new RegExp(vz,'i'); }catch(e){ reZ=null; }
      try{ if(vf) reF=new RegExp(vf,'i'); }catch(e){ reF=null; }
    }

    var totalShown=0, cCommit=0, cLocal=0, cDeleted=0, cCertif=0, cAuto=0;

    // pour chaque groupe
    $all('.group').forEach(function(g){
      var tbl=g.querySelector('.group-table');
      var rows=tbl.tBodies[0].rows;
      var shownInGroup=0;

      for(var i=0;i<rows.length;i++){
        var r=rows[i];
        var zoneday=(r.getAttribute('data-zoneday')||'').toString();
        var type=r.getAttribute('data-type')||r.cells[1].innerText.toLowerCase();
        var zoneid=(r.getAttribute('data-zoneid')||r.cells[2].innerText).toLowerCase();
        var zclass=r.getAttribute('data-class')||r.cells[3].innerText.toLowerCase();
        var fpath=(r.getAttribute('data-fpath')||'').toLowerCase();

        var vis=true;
        if(vz){ vis = useRegex && reZ ? reZ.test(zoneid) : zoneid.indexOf(vz)!==-1; }
        if(vis && vf){ vis = useRegex && reF ? reF.test(fpath) : fpath.indexOf(vf)!==-1; }
        if(vis && ty){ vis = (type===ty); }
        if(vis && cl){ vis = (zclass===cl); }
        if(vis && todayOnly){ vis = (zoneday===dy); }

        r.style.display=vis?'':'none';
        if(vis){
          shownInGroup++; totalShown++;
          if(type==='commit') cCommit++; else if(type==='local') cLocal++; else if(type==='deleted') cDeleted++;
          if(zclass==='certif') cCertif++; else cAuto++;
        }
      }

      // maj badge du groupe + masquer groupe si vide
      g.style.display = shownInGroup>0 ? '' : 'none';
      var badge = g.querySelector('.g-count'); if(badge) badge.textContent = shownInGroup;
    });

    $('count').innerText=totalShown;
    $('bCommit').innerText=cCommit;
    $('bLocal').innerText=cLocal;
    $('bDeleted').innerText=cDeleted;
    $('bCertif').innerText=cCertif;
    $('bAuto').innerText=cAuto;
    saveState();
  }

  function resetFilters(){
    $('fz').value=''; $('ff').value=''; $('ft').value=''; $('fc').value='';
    $('todayOnly').checked=false; $('regex').checked=false;
    $('colPath').checked=true; $('colHash').checked=true; $('colCHash').checked=true; $('colStart').checked=true; $('colEnd').checked=true;
    document.body.classList.remove('dark'); $('dark').checked=false;
    filterRows(); applyColVisibility();
  }

  function copyPath(p){ try{navigator.clipboard.writeText(p);}catch(e){} }

  function downloadFiltered(){
    var vis=[true,true,true,true,$('colPath').checked,$('colHash').checked,$('colCHash').checked,true,$('colStart').checked,$('colEnd').checked,true];
    var headers=[];
    // on prend l'entête du premier tableau visible
    var firstTable = document.querySelector('.group-table');
    for(var i=0;i<firstTable.tHead.rows[0].cells.length;i++){
      if(!vis[i]) continue;
      headers.push(firstTable.tHead.rows[0].cells[i].innerText);
    }
    var lines=[headers.join(';')];

    document.querySelectorAll('.group-table').forEach(function(tbl){
      var rows=tbl.tBodies[0].rows;
      for(var i=0;i<rows.length;i++){
        if(rows[i].style.display==='none') continue;
        var cols=[], cells=rows[i].cells;
        for(var j=0;j<cells.length;j++){
          if(!vis[j]) continue;
          var txt=cells[j].innerText.replace(/(\r\n|\n|\r)/g,' ').replace(/;/g,',');
          cols.push(txt);
        }
        lines.push(cols.join(';'));
      }
    });

    var blob=new Blob([lines.join('\n')],{type:'text/csv;charset=utf-8;'});
    var url=URL.createObjectURL(blob);
    var a=document.createElement('a'); a.href=url; a.download='report_grouped_'+DATEYMD+'.csv'; a.click();
    setTimeout(function(){ URL.revokeObjectURL(url); }, 1000);
  }

  window.addEventListener('DOMContentLoaded',function(){
    if(AUTO_ZONES!=='ON'){ if(PREF_CERTIF){ $('fc').value='certif'; } }
    applyColVisibility(); filterRows();

    $('fz').addEventListener('keyup',debounce(filterRows,150));
    $('ff').addEventListener('keyup',debounce(filterRows,150));
    $('ft').addEventListener('change',filterRows);
    $('fc').addEventListener('change',filterRows);
    $('todayOnly').addEventListener('change',filterRows);
    $('regex').addEventListener('change',filterRows);

    $('colPath').addEventListener('change',function(){applyColVisibility();saveState()});
    $('colHash').addEventListener('change',function(){applyColVisibility();saveState()});
    $('colCHash').addEventListener('change',function(){applyColVisibility();saveState()});
    $('colStart').addEventListener('change',function(){applyColVisibility();saveState()});
    $('colEnd').addEventListener('change',function(){applyColVisibility();saveState()});
    $('dark').addEventListener('change',function(){ document.body.classList.toggle('dark',this.checked); saveState(); });

    $('btnReset').addEventListener('click',resetFilters);
    $('btnExport').addEventListener('click',downloadFiltered);
    loadState(); filterRows();
  });

  window.copyPath=copyPath;
})();
</script>
</head>
<body>
<div class="header">
  <div class="header-inner">
    <div class="h-title">Rapport changements (group&#233; par fichier)</div>
    <label><input type="checkbox" id="dark"> Mode sombre</label>
    <span class="badge">Run: __DATEYMD__</span>
    <span class="badge">Auto zones: __AUTO_ZONES__</span>
  </div>
</div>

<details class="help">
  <summary>Aide</summary>
  <div style="margin-top:8px; line-height:1.5">
    <p><b>Id&#233;e simple</b> : un panneau par fichier. Clique sur la barre pour replier/d&#233;plier. Les filtres s&#39;appliquent &#224; tous les panneaux, et un panneau sans r&#233;sultats est masqu&#233;.</p>
    <p><b>Concret</b> : chaque panneau contient un tableau triable (par colonne) pour ce fichier. Les compteurs globaux (commit/local/etc.) ne comptent que ce qui est visible.</p>
  </div>
</details>

<div class="toolbar">
  <label>zoneid <input id="fz" type="text" placeholder="ex: GEN_1 (texte ou regex)"></label>
  <label>filepath <input id="ff" type="text" placeholder="ex: myfile.xcce (texte ou regex)"></label>
  <label>type
    <select id="ft">
      <option value="">Tous</option><option value="commit">commit</option><option value="local">local</option><option value="deleted">deleted</option>
    </select>
  </label>
  <label>classe
    <select id="fc"><!-- options injectées --></select>
  </label>
  <label><input type="checkbox" id="todayOnly"> Limiter aux modifs du jour (__DATEYMD__)</label>
  <label><input type="checkbox" id="regex"> Regex</label>
  <span>Affich&#233;s: <span id="count" class="badge">0</span></span>
  <span>commit <span id="bCommit" class="badge">0</span></span>
  <span>local <span id="bLocal" class="badge">0</span></span>
  <span>deleted <span id="bDeleted" class="badge">0</span></span>
  <span>CERTIF <span id="bCertif" class="badge">0</span></span>
  <span>auto <span id="bAuto" class="badge">0</span></span>
  <button id="btnExport">Exporter CSV (filtr&#233;)</button>
  <button id="btnReset">R&#233;initialiser</button>
  <span class="opt-cols" style="margin-left:auto">
    Colonnes:
    <label><input type="checkbox" id="colPath" checked> path</label>
    <label><input type="checkbox" id="colHash" checked> pathhash</label>
    <label><input type="checkbox" id="colCHash" checked> contenthash</label>
    <label><input type="checkbox" id="colStart" checked> start</label>
    <label><input type="checkbox" id="colEnd" checked> end</label>
  </span>
</div>
'@

  # injecte variables
  $header = $header.Replace('__DATEYMD__', $DateYmd).Replace('__AUTO_ZONES__', $AUTO_ZONES).Replace('__PREF_CERTIF__', $selExpJs)

  $html  = New-Object System.Text.StringBuilder
  $null = $html.AppendLine($header)

  # options "classe" (UI: 'certif', 'auto_step', 'auto_file')
  $selectedCertif = if($null -ne $hasCertif){ ' selected' } else { '' }
  if($AUTO_ZONES -eq 'ON'){
    $classOptions = '<option value="">Toutes</option>' +
                    '<option value="certif"'+ $selectedCertif +'>CERTIF</option>' +
                    '<option value="auto_step">AUTO: Step</option>' +
                    '<option value="auto_file">AUTO: File</option>'
  } else {
    $classOptions = '<option value="certif"'+ $selectedCertif +'>CERTIF</option>'
  }
  $null = $html.AppendLine('<script>document.addEventListener("DOMContentLoaded",function(){ document.getElementById("fc").innerHTML = ' + '"' + ($classOptions -replace '"','\"') + '"' + '; });</script>')

  # ===== rendu des groupes =====
  $gid = 0
  foreach($g in $groups){
    $gid++
    $tblId = "tbl_$gid"
    $fpEncGroup = $g.Name -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'

    $null = $html.AppendLine('<details class="group" open id="grp_'+$gid+'">')
    $null = $html.AppendLine('<summary><span class="g-title">'+$fpEncGroup+'</span><span class="g-badges"><span class="badge">lignes visibles: <span class="g-count">0</span></span></span></summary>')

    # table du groupe
    $null = $html.AppendLine('<div class="g-body"><table class="group-table" id="'+$tblId+'"><thead><tr>' +
      '<th onclick="sg('''+$tblId+''',0)">modifi&eacute; le</th>' +
      '<th onclick="sg('''+$tblId+''',1)">type</th>' +
      '<th onclick="sg('''+$tblId+''',2)">zoneid</th>' +
      '<th onclick="sg('''+$tblId+''',3)">zoneclass</th>' +
      '<th onclick="sg('''+$tblId+''',4)">filepath</th>' +
      '<th onclick="sg('''+$tblId+''',5)">pathhash</th>' +
      '<th onclick="sg('''+$tblId+''',6)">contenthash</th>' +
      '<th onclick="sg('''+$tblId+''',7)">logfile</th>' +
      '<th onclick="sg('''+$tblId+''',8)">start</th>' +
      '<th onclick="sg('''+$tblId+''',9)">end</th>' +
      '<th>actions</th>' +
    '</tr></thead><tbody>')

    foreach($r in ($g.Group | Sort-Object zone_ts -Descending)){
      $fpEnc = $r.filepath -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
      $lfEnc = $r.logfile  -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
      $href  = 'file:///' + ($r.logfile -replace '\\','/')
      $hrefEnc = $href -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
      $dataPath = $r.logfile -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'

      $zclassDisplay = if($r.zoneclass -eq 'explicit'){ 'CERTIF' } else { $r.zoneclass }
      $zclassCss     = if($r.zoneclass -eq 'explicit'){ 'certif' } else { $r.zoneclass }

      $zoneDay = try { [datetime]::Parse($r.zone_ts,[System.Globalization.CultureInfo]::InvariantCulture).ToString('yyyyMMdd') } catch { $DateYmd }

      $rowClass = 'type-' + $r.type + ' zone-' + $zclassCss
      # on duplique des infos en data-attrs pour des filtres rapides
      $line = '<tr class="'+ $rowClass +'" data-zoneday="'+ $zoneDay +'" data-type="'+ $r.type.ToLower() +'" data-zoneid="'+ ($r.zoneid -replace '"','&quot;') +'" data-class="'+ $zclassCss +'" data-fpath="'+ ($r.filepath.ToLower() -replace '"','&quot;') +'">' +
              '<td>'+ $r.zone_ts +'</td><td>'+ $r.type +'</td><td>'+ $r.zoneid +'</td><td>'+ $zclassDisplay +'</td><td>'+ $fpEnc +'</td><td>'+ $r.pathhash +'</td><td>'+ $r.contenthash +'</td><td><a href="'+ $hrefEnc +'" target="_blank" rel="noopener">'+ $lfEnc +'</a></td><td>'+ $r.start +'</td><td>'+ $r.end +'</td><td><button onclick="copyPath(&quot;'+ $dataPath +'&quot;)">Copier chemin log</button></td></tr>'
      $null = $html.AppendLine($line)
    }

    $null = $html.AppendLine('</tbody></table></div></details>')
  }

  $null = $html.AppendLine('</body></html>')

  $out = Join-Path $OutDir ("report_{0}.html" -f $DateYmd)
  [IO.File]::WriteAllText($out,$html.ToString(),[Text.Encoding]::UTF8)
}


Build-HtmlReport -Csv $CSV_FILE -OutDir $LOG_DIR -DateYmd $DATE_YMD
Write-Info ("Summary : '{0}'" -f $summaryPath)
Write-Info ("HTML    : '{0}'" -f (Join-Path $LOG_DIR ("report_{0}.html" -f $DATE_YMD)))
Write-Info 'Fin.'
