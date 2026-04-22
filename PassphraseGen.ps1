Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
# Microsoft.Win32.SaveFileDialog lives in PresentationFramework - no separate Add-Type needed

# ==================================================================
#  SINGLE INSTANCE  - only one copy can run at a time
# ==================================================================
$mutexName = "PassphraseGenerator_SingleInstance"
$mutex = New-Object System.Threading.Mutex($false, $mutexName)
$owned = $false
try {
    $owned = $mutex.WaitOne(0, $false)
} catch [System.Threading.AbandonedMutexException] {
    $owned = $true
}
if (-not $owned) {
    # Find and bring the existing window to the foreground
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
    $proc = Get-Process | Where-Object { $_.MainWindowTitle -eq "Passphrase Generator" } | Select-Object -First 1
    if ($proc) {
        [Win32]::ShowWindow($proc.MainWindowHandle, 9) | Out-Null
        [Win32]::SetForegroundWindow($proc.MainWindowHandle) | Out-Null
    }
    exit
}

# ==================================================================
#  HIDE THE POWERSHELL CONSOLE WINDOW
# ==================================================================
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class ConsoleHider {
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]   public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
$consoleHandle = [ConsoleHider]::GetConsoleWindow()
[ConsoleHider]::ShowWindow($consoleHandle, 0) | Out-Null  # 0 = SW_HIDE

# ==================================================================
#  DETECT WINDOWS THEME  (light = 1, dark = 0)
# ==================================================================
function Get-WindowsTheme {
    try {
        $val = Get-ItemPropertyValue `
            -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" `
            -Name "AppsUseLightTheme" -ErrorAction Stop
        return if ($val -eq 1) { "WindowsLight" } else { "WindowsDark" }
    } catch { return "WindowsDark" }
}

# ==================================================================
#  THEME DEFINITIONS
# ==================================================================
$THEMES = @{

    # -- Windows Dark ----------------------------------------------
    WindowsDark = @{
        WinBg         = "#1c1c1c"
        PanelBg       = "#2d2d2d"
        PanelBorder   = "#3d3d3d"
        OutputBg      = "#252525"
        OutputBorder  = "#404040"
        BarTrack      = "#3a3a3a"
        BarFill       = "#0078d4"
        HeaderLabel   = "#0078d4"
        HeaderSub     = "#f3f3f3"
        OutputText    = "#d4eeff"
        StatLabel     = "#707070"
        StatValue     = "#0078d4"
        SliderFill    = "#0078d4"
        SliderTrack   = "#3a3a3a"
        ThumbStroke   = "#60b0ff"
        CBBorder      = "#404040"
        CBCheck       = "#0078d4"
        CBBg          = "#2d2d2d"
        ComboText     = "#c8c8c8"
        ComboBg       = "#2d2d2d"
        ComboBorder   = "#404040"
        CopyFg        = "#60b0ff"
        CopyBg        = "#1a2a3a"
        CopyBorder    = "#0078d4"
        GenBg         = "#0078d4"
        GenBorder     = "#60b0ff"
        GenFg         = "#ffffff"
        DividerColor  = "#3d3d3d"
        MutedText     = "#707070"
        CountFg       = "#0078d4"
    }

    # -- Windows Light ---------------------------------------------
    WindowsLight = @{
        WinBg         = "#f3f3f3"
        PanelBg       = "#ffffff"
        PanelBorder   = "#d0d0d0"
        OutputBg      = "#fafafa"
        OutputBorder  = "#c8c8c8"
        BarTrack      = "#d8d8d8"
        BarFill       = "#0078d4"
        HeaderLabel   = "#0063b1"
        HeaderSub     = "#1a1a1a"
        OutputText    = "#003a6e"
        StatLabel     = "#767676"
        StatValue     = "#0063b1"
        SliderFill    = "#0078d4"
        SliderTrack   = "#d0d0d0"
        ThumbStroke   = "#0078d4"
        CBBorder      = "#c0c0c0"
        CBCheck       = "#0078d4"
        CBBg          = "#ffffff"
        ComboText     = "#1a1a1a"
        ComboBg       = "#ffffff"
        ComboBorder   = "#c0c0c0"
        CopyFg        = "#0063b1"
        CopyBg        = "#e8f2fb"
        CopyBorder    = "#90bde0"
        GenBg         = "#0078d4"
        GenBorder     = "#0063b1"
        GenFg         = "#ffffff"
        DividerColor  = "#d0d0d0"
        MutedText     = "#767676"
        CountFg       = "#0063b1"
    }

    # -- KPop Demon Hunters ----------------------------------------
    # Deep midnight base  .  magenta demon glow  .  iridescent gold
    # candy pink/lavender/cyan accents  .  soft K-drama warmth
    KPopDemon = @{
        WinBg         = "#0d0714"      # near-black with a purple undertone
        PanelBg       = "#160d22"      # deep violet-dark
        PanelBorder   = "#3d1f5c"      # muted purple border
        OutputBg      = "#120a1e"
        OutputBorder  = "#6b1f7a"      # demon magenta border
        BarTrack      = "#1e1030"
        BarFill       = "#e040a0"      # hot magenta (Saja Boys energy)
        HeaderLabel   = "#f07acd"      # candy pink label
        HeaderSub     = "#f0e6ff"      # soft lavender-white
        OutputText    = "#ffd6f5"      # pastel pink passphrase text
        StatLabel     = "#5a3a72"
        StatValue     = "#f07acd"
        SliderFill    = "#c850e8"      # vivid purple-magenta
        SliderTrack   = "#1e1030"
        ThumbStroke   = "#ffd6f5"
        CBBorder      = "#3d1f5c"
        CBCheck       = "#e040a0"
        CBBg          = "#160d22"
        ComboText     = "#d8b4fe"      # soft lilac
        ComboBg       = "#160d22"
        ComboBorder   = "#3d1f5c"
        CopyFg        = "#f0abdc"
        CopyBg        = "#1e0a2e"
        CopyBorder    = "#6b1f7a"
        GenBg         = "#8b0fa8"      # deep conjure-purple
        GenBorder     = "#e040a0"
        GenFg         = "#ffeef8"
        DividerColor  = "#2e1445"
        MutedText     = "#5a3a72"
        CountFg       = "#f07acd"
    }
}

# ==================================================================
#  WORDLIST  - EFF Large Diceware list (7776 words, ~12.9 bits/word)
#  Download: https://www.eff.org/files/2016/07/18/eff_large_wordlist.txt
#  Place the file in the same folder as this script:
#    C:\Temp\PasswordGen\eff_large_wordlist.txt
# ==================================================================
$WORDLIST_PATH = Join-Path $PSScriptRoot "eff_large_wordlist.txt"

if (Test-Path $WORDLIST_PATH) {
    # Parse the EFF file: each line is "DDDDD<tab>word" - grab just the word
    $WORDLIST = (Get-Content $WORDLIST_PATH) |
        Where-Object  { $_ -match '^\d' } |
        ForEach-Object { ($_ -split '\s+')[1].Trim() } |
        Where-Object  { $_ -ne '' }
    Write-Host "Loaded EFF large wordlist: $($WORDLIST.Count) words (~$([Math]::Round([Math]::Log($WORDLIST.Count,2),1)) bits/word)"
} else {
    # -- Fallback: built-in 1296-word list (~10.3 bits/word) ------
    Write-Host "EFF wordlist not found at '$WORDLIST_PATH' - using built-in fallback list."
    Write-Host "For maximum security, download the EFF large wordlist from:"
    Write-Host "  https://www.eff.org/files/2016/07/18/eff_large_wordlist.txt"
    Write-Host "and save it to: $WORDLIST_PATH"
    $WORDLIST = @(
"aardvark","ability","abnormal","abrasive","absorbing","abundant","abyss","academic","accurate","achieve",
"acoustic","acquire","activate","adamant","addition","adequate","adjacent","adorable","advanced","adverse",
"aerial","affirm","agenda","agility","agnostic","agonize","agreeable","aircraft","algebra","algorithm",
"alienate","aligned","alligator","allocate","allowance","alongside","altitude","aluminum","amalgam","ambiguity",
"ambitious","amplified","amicable","amusing","analogous","analyst","anarchy","ancestor","ancient","angular",
"animated","announce","anthology","antidote","anvil","apparent","appetite","applause","applied","approach",
"aqueduct","arcade","archival","arduous","aroma","artefact","artistic","artwork","asphalt","assembled",
"asteroid","athletic","atmospheric","atrocity","attached","audacious","authentic","authority","automatic","avalanche",
"aviation","avoidance","backdrop","backfire","backbone","badminton","balanced","barricade","basement","bathroom",
"battalion","bearing","beautiful","becoming","behavior","benchmark","benevolent","bewilder","bicycling","bilateral",
"biography","blanket","blast","blizzard","blockade","blueprint","boardroom","bolster","boundary","bravery",
"breakout","brilliant","brittle","broadband","building","bulletin","bureaucrat","bustling","calendar","campaign",
"canister","canyon","capacity","capture","cathedral","cavalier","cavern","cellular","ceremony","challenge",
"champion","channel","charisma","chemical","circuit","citation","civilian","classic","climate","coalition",
"cognitive","coherent","collapse","collide","colossal","combat","commerce","complexity","concrete","confident",
"construct","continent","contrast","coordinate","corridor","courageous","coverage","creative","criminal","critical",
"crowded","cryptic","currency","cylinder","dangerous","database","daylight","dazzling","deadlock","debris",
"decisive","dedicated","defense","dehydrate","delirious","delivery","demanding","democracy","demolish","dependent",
"describe","detailed","detective","diagonal","diameter","diplomat","discover","disorder","dispatch","district",
"dominant","dramatic","durable","dynamic","elaborate","eliminate","emergent","emphasis","endurance","enforce",
"enhance","enormous","epidemic","evasive","evolving","excellent","exchange","exciting","exclusive","exhausted",
"explicit","explosive","exponent","exterior","extreme","fabulous","fairness","familiar","ferocious","festival",
"firestorm","flexible","folklore","forecast","forested","fracture","fragment","framework","freedom","frontline",
"function","furnace","futuristic","generate","glorious","gradient","granite","grateful","gridlock","guardian",
"guidance","gymnasium","habitat","hallmark","hardship","harmony","harvest","hazardous","headline","heroic",
"hierarchy","highland","historical","holistic","horizon","hydraulic","ideology","illuminate","immovable","imperial",
"imposing","incognito","increment","inertia","infinite","informed","inherent","innocent","install","integral",
"intense","internal","intricate","invasion","invincible","isolated","judicial","junction","keystroke","labyrinth",
"landmark","language","leverage","lightning","lionhearted","latitude","logical","luminous","machinery","manifest",
"martial","massive","masterful","mechanical","mercenary","momentum","mountain","mutable","narrative","national",
"navigate","network","neutral","nitrogen","notable","nuclear","obstacle","offshore","operative","orderly",
"outpost","overcome","panoramic","parallel","parameter","partisan","passage","pathway","patriotic","perimeter",
"phantom","platform","powerful","practise","precision","premium","pressure","primary","priority","pristine",
"proactive","profound","prolific","prominent","protocol","province","pursuit","qualified","quantity","radiant",
"rational","reaction","realm","recommend","reliable","resistant","resolve","retrieve","rigorous","robust",
"runaway","safeguard","satellite","seamless","sectoral","sentinel","severity","signal","singular","skeleton",
"slippery","snapshot","solution","sovereign","spectrum","stability","stamina","standard","strategy","structure",
"supreme","tactical","temperate","terminal","texture","thorough","threshold","titanium","together","tornado",
"trackable","tradition","trajectory","turbulent","ultimate","unanimous","undaunted","uniform","unstable","validate",
"velocity","verdict","vigilant","volatile","wavelength","wilderness","workforce","zodiac","absolute","adhesive",
"ambience","antennae","artifact","aversion","battery","borealis","capable","carbonate","cassette","cautious",
"centrist","chatroom","chromium","circular","clearcut","cliffside","clockwork","cobblestone","coldfront","colorfast",
"compress","concealed","conflicts","continuum","coolant","corrected","countable","covenant","cracking","crimson",
"crossbow","cupboard","deadline","decimals","declared","defended","deletion","departed","derailed","distress",
"dormant","drafting","drilling","dropzone","eastward","economy","elevated","embedded","encoding","endpoint",
"enforced","engaging","enlarged","enlisted","entitled","entrance","equation","eruption","evasion","excluded",
"existing","expanded","exported","fastened","filtered","finalized","focused","followed","framed","gathered",
"governed","hardened","imported","indexed","inferred","launched","layered","learned","limited","managed",
"measured","migrated","modified","mounted","ordered","patched","planned","plotted","pointed","powered",
"prepared","preserved","processed","produced","programmed","projected","promoted","proven","queried","rebuilt",
"received","recorded","reduced","released","remained","removed","replaced","resolved","restored","retained",
"returned","reviewed","revised","rotated","routed","sampled","scanned","secured","selected","serviced",
"shifted","signaled","simulated","sourced","started","stored","supplied","tested","tracked","trained",
"transferred","triggered","unified","unlocked","updated","upgraded","utilized","verified","watched","written",
"abstain","abysmal","accredit","acetone","acrobat","acutely","adeptly","admiral","adornment","afterglow",
"airborne","airfield","airspace","alchemy","alerting","almanac","altering","ambushed","amphibian","analysis",
"anemone","annotate","aperture","archetype","armistice","articulate","ascertain","aspiring","assailant","assurance",
"attainment","attrition","augmented","avoidable","backlight","backtrack","bandwidth","barometer","bedrock","beleaguer",
"belligerent","blindspot","bombshell","bottleneck","breakaway","broadcast","brutalize","bulwark","camouflage","captivate",
"cartridge","cataclysm","catapult","celestial","centrifuge","chainmail","checkmate","ciphered","clearance","clipboard",
"cloudbank","codebreak","coldstart","collision","commander","commando","commodity","compressor","concentric","conductor",
"cornered","counteract","crackling","crossfire","crosstalk","cryptogram","cutthroat","datastream","deadzone","debrief",
"decrypted","defector","detonator","deviant","dismantle","disruptor","divulged","doctrine","dominance","downlink",
"dreadnought","encrypted","enforcer","escalate","estimator","exfiltrate","exosphere","expedited","extracted","firsthand",
"fissure","flashpoint","floodgate","flywheel","footprint","freefall","frequency","fulcrum","geothermal","glimmer",
"governess","gridlock","groundwork","gyrosphere","hackproof","halogen","hammerhead","hardpoint","hawkseye","headstart",
"heatshield","heavyload","heliodor","heliostat","hellfire","highpoint","hightide","hotswap","hurricane","hyperdrive",
"hyperlink","icebreaker","immersion","inferno","informant","insolvent","intercept","ironvault","jackknife","jamboree",
"javelin","jumpstart","kilowatt","knapsack","launcher","layover","leadoff","loadout","lockstep","lodestone",
"longshot","mainboard","maingate","mainline","mandate","marauder","masquerade","matchpoint","meltdown","midpoint",
"milestone","mineswept","missionlog","mitigation","moonstone","motivate","neutralize","nightfall","nightwatch","nominal",
"northstar","nullified","objective","obstruct","onramp","openfire","ordnance","outbreak","outreach","pacesetter",
"paramount","peacemaker","phaseline","pincer","pitfall","pointblank","polished","powercore","powerdown","prestige",
"primeline","proofread","prowler","purgatory","quicksand","radiogram","raincloud","rampant","rebound","reclaim",
"recoil","recovery","redirect","redundant","regiment","relentless","remapped","retaliate","ricochet","rocketfuel",
"sabotage","safezone","sandcastle","scattershot","scoreboard","scrambled","searchlight","seismic","shadowrun","shockwave",
"sidearm","sightline","skirmish","skywatch","slipstream","smokeless","snakepit","sonarping","sophistry","sortie",
"soundwall","southpaw","spacegate","speedrun","spotcheck","stampede","standfast","steadfast","stealth","stormbolt",
"stormfront","strikezone","stronghold","submerged","surveyor","swiftness","switchback","swordfish","taskmaster",
"tempestuous","throttled","thunderhead","tidewater","timeshift","titanfall","torchlight","touchpoint","traceback","trapdoor",
"trenched","tripwire","turbojet","turnpike","unarmored","unchecked","underfire","undetected","undertaker","unguarded",
"unpatched","unshielded","uplinked","uprooted","vaultdoor","vortex","warcry","warden","warfront","warpzone",
"watchdog","watchpost","waterfall","wildfire","windhover","withheld","wolfpack","workbench","worldline","wreckage",
"zenith","zerohour","airdrop","alluvial","amplifier","anchorage","anomalous","appended","aquifer","arborist",
"armament","ascendant","assembly","backslash","basecamp","biometric","birthmark","blackout","boredom","calmness",
"canopy","castaway","catalyst","cavitate","centenary","champion","chartroom","checklist","chieftain","chokehold",
"cipherkey","clearway","codename","combative","commuter","condenser","confident","conquest","contender","cooldown",
"copywrite","corrosion","crosshair","culprit","cyberspace","darkroom","daybreak","decrypt","defiance","delegate",
"derelict","destruct","detected","deviation","diffused","digitized","director","displace","dropout","earthwork",
"echelon","edgecase","egress","embattle","emissary","empowered","enclave","endpoint","engineer","entrenched",
"equinox","eradicate","escalator","ethanol","eviction","exchange","exhaust","exosphere","expanse","expedient",
"exploited","extractor","falsehood","fastbreak","faultline","fencepost","ferocity","firmware","fissured","flagship",
"flankguard","floodwall","flyover","foothold","fortress","freeform","frontier","frostbite","fullscale","fuselage",
"gameplan","gauntlet","gearshift","generator","glassceil","glidepath","globespan","grappling","greenzone","gridmap",
"groundzero","guardrail","gunpoint","halflife","hardline","heatwave","hexagonal","highrise","holdpoint","hotline",
"icewall","ignition","impasse","implanted","impulse","incoming","infiltrate","infrared","insertion","ironside",
"jailbreak","jetstream","junction","killjoy","knockout","landfall","landslide","launchpad","layermap","leapfrog",
"lockpick","lodestar","lowlands","mainmast","maneuvre","marksman","maxrange","megabyte","meridian","methane",
"midfield","minefield","mineshaft","missiles","modulate","nightside","northward","notebook","nullzone","offshore",
"openrange","outflank","outpost","overhaul","overlord","overwatch","oxidized","paragon","partisan","peacefire",
"periscope","piercing","pinpointed","pipeline","pitchblack","plaintext","playbook","plunder","podium","pointman",
"powerline","precision","prepstage","primetime","printout","profile","propellant","proton","pushback","quarantine",
"quickdraw","quickstep","rearguard","receptor","redacted","redirect","redzone","regroup","reinforce","remnant",
"remotely","rendezvous","repeater","resolver","restrike","retrofit","rewind","ridgeline","riskzone","roadblock",
"roughage","safehouse","salvo","scanner","scarlet","scramble","seawall","sectioned","sendoff","sentinel",
"shakedown","sidequest","skydive","sleeper","slideshow","sniperhole","snowline","southward","spearhead","splinter",
"spotlight","standoff","stormwall","stratagem","streamline","subbasement","subgraph","substation","sundial","supercell",
"swampland","sweepfire","switchfire","syndicate","tailwind","takedown","teleport","terminal","testbed","thinline",
"throwback","thunderbolt","tightrope","timeline","toehold","topside","toughline","traceroute","trailblazer","transmit",
"trapline","treeline","tribunal","tundra","turbulence","twilight","twinfire","ultrasonic","undercover","undermine",
"underpass","undertone","unevenly","unified","uprising","upshot","uranium","vanguard","variable","vaultline",
"ventilate","voidspace","voltage","vortexmap","warfare","warzone","wasteland","waveform","waypoint","webwork",
"wellpoint","westward","wingspan","wiretap","withdraw","wormgate","yearlong","yieldpoint","zeroloss","zipline"
    )
}

# ==================================================================
#  CRYPTO RNG  (rejection sampling - no modulo bias)
# ==================================================================
function Get-CryptoRandom {
    param([int]$max)
    $rng   = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $bytes = New-Object byte[] 4
    $limit = [uint32]::MaxValue - ([uint32]::MaxValue % [uint32]$max)
    do {
        $rng.GetBytes($bytes)
        $val = [BitConverter]::ToUInt32($bytes, 0)
    } while ($val -ge $limit)
    return [int]($val % $max)
}

function New-Passphrase {
    param([int]$wordCount, [string]$separator)

    # Build word list
    $words = [System.Collections.Generic.List[string]]::new()
    for ($i = 0; $i -lt $wordCount; $i++) {
        $w = $WORDLIST[(Get-CryptoRandom -max $WORDLIST.Count)]
        $words.Add($w.Substring(0,1).ToUpper() + $w.Substring(1).ToLower())
    }

    # Pick one random word index and append a crypto-random 2-digit number (10-99)
    $targetIdx  = Get-CryptoRandom -max $wordCount
    $twoDigit   = (Get-CryptoRandom -max 90) + 10   # 10..99
    $words[$targetIdx] = $words[$targetIdx] + $twoDigit.ToString()

    return $words -join $separator
}

function Get-EntropyBits {
    param([int]$wordCount, [int]$listSize)
    [Math]::Round($wordCount * [Math]::Log($listSize, 2), 1)
}

function Get-CrackTime {
    param([double]$entropyBits)
    $sec = [Math]::Pow(2, $entropyBits) / 1e12
    $min = 60; $hr = 3600; $day = 86400; $yr = 31536000
    if     ($sec -lt 1)          { "less than a second" }
    elseif ($sec -lt $min)       { "$([Math]::Round($sec)) seconds" }
    elseif ($sec -lt $hr)        { "$([Math]::Round($sec/$min)) minutes" }
    elseif ($sec -lt $day)       { "$([Math]::Round($sec/$hr)) hours" }
    elseif ($sec -lt $yr)        { "$([Math]::Round($sec/$day)) days" }
    elseif ($sec -lt $yr*100)    { "$([Math]::Round($sec/$yr)) years" }
    elseif ($sec -lt $yr*1e3)    { "$([Math]::Round($sec/($yr*100))) centuries" }
    elseif ($sec -lt $yr*1e6)    { "$([Math]::Round($sec/($yr*1e3))) millennia" }
    elseif ($sec -lt $yr*1e9)    { "millions of years" }
    elseif ($sec -lt $yr*1e12)   { "billions of years" }
    else                         { "trillions of years" }
}

function Get-StrengthInfo {
    param([double]$entropy, [hashtable]$t)
    if     ($entropy -ge 128) { "FORTRESS",  "#00ff9d", 1.00 }
    elseif ($entropy -ge 100) { "EXCELLENT", "#22d3a0", 0.88 }
    elseif ($entropy -ge 80)  { "STRONG",    "#4ade80", 0.72 }
    elseif ($entropy -ge 60)  { "GOOD",      "#facc15", 0.54 }
    elseif ($entropy -ge 40)  { "FAIR",      "#fb923c", 0.36 }
    else                      { "WEAK",      "#f87171", 0.20 }
}

# ==================================================================
#  PRONOUNCEABILITY SCORE
#  Rates how easy a passphrase is to say aloud / remember.
#  Scoring per word: penalise consecutive consonants, reward
#  alternating vowel/consonant patterns and short words.
# ==================================================================
function Get-PronounceScore {
    param([string]$passphrase)
    $vowels = 'aeiouAEIOU'
    # Strip numbers/symbols, split on non-alpha
    $words = ($passphrase -split '[^a-zA-Z]+') | Where-Object { $_.Length -gt 0 }
    if ($words.Count -eq 0) { return "N/A", "#888888" }

    $totalScore = 0
    foreach ($word in $words) {
        $score    = 100
        $len      = $word.Length
        $consRun  = 0
        $prevType = ""
        $alternations = 0

        for ($i = 0; $i -lt $len; $i++) {
            $c = $word[$i]
            $isVowel = $vowels.Contains($c)
            $curType = if ($isVowel) { "V" } else { "C" }

            if ($curType -eq "C") {
                $consRun++
                if ($consRun -ge 3) { $score -= 12 }   # triple consonant cluster
                elseif ($consRun -eq 2) { $score -= 4 } # double consonant
            } else {
                $consRun = 0
            }

            if ($prevType -ne "" -and $curType -ne $prevType) { $alternations++ }
            $prevType = $curType
        }

        # Reward nice alternation ratio
        if ($len -gt 1) {
            $altRatio = $alternations / ($len - 1)
            $score += [int]($altRatio * 20)
        }

        # Penalise very long words (harder to say quickly)
        if ($len -gt 10) { $score -= ($len - 10) * 3 }

        $totalScore += [Math]::Max(0, [Math]::Min(100, $score))
    }

    $avg = [int]($totalScore / $words.Count)

    if     ($avg -ge 85) { return "Easy",   "#4ade80" }
    elseif ($avg -ge 65) { return "Good",   "#86efac" }
    elseif ($avg -ge 45) { return "Tricky", "#facc15" }
    else                 { return "Hard",   "#f87171" }
}

# ==================================================================
#  COPY SOUND  (short 880Hz beep via SoundPlayer inline WAV)
# ==================================================================
function Play-CopySound {
    try {
        Add-Type -AssemblyName System.Media
        # Generate a tiny 880Hz tone WAV in memory (0.08 sec, 44100Hz, mono 16-bit)
        $sampleRate = 44100
        $duration   = 0.08
        $freq       = 880
        $numSamples = [int]($sampleRate * $duration)
        $amplitude  = 12000

        $ms = New-Object System.IO.MemoryStream
        $bw = New-Object System.IO.BinaryWriter($ms)

        # WAV header
        $dataSize = $numSamples * 2
        $bw.Write([byte[]][System.Text.Encoding]::ASCII.GetBytes("RIFF"))
        $bw.Write([int32](36 + $dataSize))
        $bw.Write([byte[]][System.Text.Encoding]::ASCII.GetBytes("WAVEfmt "))
        $bw.Write([int32]16)           # chunk size
        $bw.Write([int16]1)            # PCM
        $bw.Write([int16]1)            # mono
        $bw.Write([int32]$sampleRate)
        $bw.Write([int32]($sampleRate * 2))
        $bw.Write([int16]2)            # block align
        $bw.Write([int16]16)           # bits per sample
        $bw.Write([byte[]][System.Text.Encoding]::ASCII.GetBytes("data"))
        $bw.Write([int32]$dataSize)

        # PCM samples - quick fade-out envelope
        for ($i = 0; $i -lt $numSamples; $i++) {
            $t       = $i / $sampleRate
            $env     = 1.0 - ($i / $numSamples)   # linear fade
            $sample  = [int]($amplitude * $env * [Math]::Sin(2 * [Math]::PI * $freq * $t))
            $bw.Write([int16]$sample)
        }
        $bw.Flush()
        $ms.Position = 0

        $player = New-Object System.Media.SoundPlayer($ms)
        $player.Play()   # async - doesn't block UI
    } catch { <# silently ignore if audio unavailable #> }
}

# ==================================================================
#  XAML
# ==================================================================
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Passphrase Generator"
    Width="700" Height="630"
    MinWidth="560" MinHeight="570"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanResizeWithGrip"
    FontFamily="Segoe UI">

  <Window.Resources>

    <!-- Rounded button template -->
    <ControlTemplate x:Key="RndBtn" TargetType="Button">
      <Border x:Name="bd"
              Background="{TemplateBinding Background}"
              BorderBrush="{TemplateBinding BorderBrush}"
              BorderThickness="{TemplateBinding BorderThickness}"
              CornerRadius="8" Padding="{TemplateBinding Padding}">
        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
      </Border>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter TargetName="bd" Property="Opacity" Value="0.78"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True">
          <Setter TargetName="bd" Property="Opacity" Value="0.52"/>
        </Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>

    <Style x:Key="GhostBtn" TargetType="Button">
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="FontSize"        Value="12"/>
      <Setter Property="FontWeight"      Value="SemiBold"/>
      <Setter Property="Padding"         Value="0"/>
      <Setter Property="Template"        Value="{StaticResource RndBtn}"/>
    </Style>

    <Style x:Key="GenBtn" TargetType="Button">
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="FontSize"        Value="14"/>
      <Setter Property="FontWeight"      Value="SemiBold"/>
      <Setter Property="Padding"         Value="0"/>
      <Setter Property="Template"        Value="{StaticResource RndBtn}"/>
    </Style>

    <Style x:Key="CopyBtn" TargetType="Button">
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="FontSize"        Value="12"/>
      <Setter Property="Padding"         Value="14,8"/>
      <Setter Property="Template"        Value="{StaticResource RndBtn}"/>
    </Style>

    <Style x:Key="ThemeBtn" TargetType="Button">
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="FontSize"        Value="11"/>
      <Setter Property="Padding"         Value="10,5"/>
      <Setter Property="Template"        Value="{StaticResource RndBtn}"/>
    </Style>

    <!-- Minimal slider -->
    <Style TargetType="Slider">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Slider">
            <Grid Height="24" VerticalAlignment="Center">
              <Track x:Name="PART_Track">
                <Track.DecreaseRepeatButton>
                  <RepeatButton Command="Slider.DecreaseLarge" Focusable="False">
                    <RepeatButton.Template>
                      <ControlTemplate TargetType="RepeatButton">
                        <Border x:Name="fill" Height="5" CornerRadius="2.5"
                                Background="#0078d4"/>
                      </ControlTemplate>
                    </RepeatButton.Template>
                  </RepeatButton>
                </Track.DecreaseRepeatButton>
                <Track.IncreaseRepeatButton>
                  <RepeatButton Command="Slider.IncreaseLarge" Focusable="False">
                    <RepeatButton.Template>
                      <ControlTemplate TargetType="RepeatButton">
                        <Border x:Name="track" Height="5" CornerRadius="2.5"
                                Background="#3a3a3a"/>
                      </ControlTemplate>
                    </RepeatButton.Template>
                  </RepeatButton>
                </Track.IncreaseRepeatButton>
                <Track.Thumb>
                  <Thumb>
                    <Thumb.Template>
                      <ControlTemplate TargetType="Thumb">
                        <Ellipse x:Name="thumb" Width="18" Height="18"
                                 Fill="#0078d4" Stroke="#60b0ff" StrokeThickness="2"/>
                      </ControlTemplate>
                    </Thumb.Template>
                  </Thumb>
                </Track.Thumb>
              </Track>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- CheckBox -->
    <Style TargetType="CheckBox">
      <Setter Property="FontSize"   Value="13"/>
      <Setter Property="Cursor"     Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
              <Border x:Name="box" Width="18" Height="18" CornerRadius="4"
                      BorderThickness="1.5" Margin="0,0,8,0" VerticalAlignment="Center">
                <TextBlock x:Name="chk" Text="v" FontSize="11" FontWeight="Bold"
                           HorizontalAlignment="Center" VerticalAlignment="Center"
                           Visibility="Collapsed"/>
              </Border>
              <ContentPresenter VerticalAlignment="Center"/>
            </StackPanel>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="chk" Property="Visibility" Value="Visible"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ComboBox -->
    <Style TargetType="ComboBox">
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding"         Value="10,7"/>
      <Setter Property="FontSize"        Value="13"/>
      <Setter Property="Height"          Value="38"/>
    </Style>

  </Window.Resources>

  <!-- Root -->
  <Border x:Name="rootBorder">
    <Grid Margin="28,16,28,22">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>   <!-- 0  Theme bar -->
        <RowDefinition Height="12"/>
        <RowDefinition Height="Auto"/>   <!-- 2  Header -->
        <RowDefinition Height="14"/>
        <RowDefinition Height="Auto"/>   <!-- 4  Output -->
        <RowDefinition Height="10"/>
        <RowDefinition Height="8"/>      <!-- 6  Strength bar -->
        <RowDefinition Height="10"/>
        <RowDefinition Height="Auto"/>   <!-- 8  Stats row -->
        <RowDefinition Height="14"/>
        <RowDefinition Height="Auto"/>   <!-- 10 Settings -->
        <RowDefinition Height="14"/>
        <RowDefinition Height="46"/>     <!-- 12 Generate -->
      </Grid.RowDefinitions>

      <!-- == THEME BAR == -->
      <StackPanel Grid.Row="0" Orientation="Horizontal" HorizontalAlignment="Right">
        <TextBlock x:Name="lblThemeHead" Text="THEME" FontSize="9" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,10,0"/>
        <Button x:Name="btnThemeWin"  Content="Windows"    Style="{StaticResource ThemeBtn}" Margin="0,0,6,0"/>
        <Button x:Name="btnThemeKpop" Content="KPDH"     Style="{StaticResource ThemeBtn}"/>
      </StackPanel>

      <!-- == HEADER == -->
      <Grid Grid.Row="2">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel>
          <TextBlock x:Name="lblHeaderTop" Text="CRYPTOGRAPHIC" FontSize="10"
                     FontWeight="Bold"/>
          <TextBlock x:Name="lblHeaderSub" Text="Passphrase Generator"
                     FontSize="24" FontWeight="Light"/>
        </StackPanel>
        <TextBlock x:Name="lblCharLen" Grid.Column="1"
                   VerticalAlignment="Bottom" FontSize="11" Text=""/>
      </Grid>

      <!-- == OUTPUT == -->
      <Border x:Name="outputBorder" Grid.Row="4"
              BorderThickness="1" CornerRadius="10" Padding="18,14">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <TextBox x:Name="txtPassphrase"
                   Background="Transparent" BorderThickness="0"
                   FontFamily="Cascadia Mono,Consolas,Courier New,monospace"
                   FontSize="18" FontWeight="Medium"
                   IsReadOnly="True" TextWrapping="Wrap"
                   VerticalScrollBarVisibility="Auto"
                   MinHeight="52" MaxHeight="100"
                   VerticalAlignment="Center"
                   Text="Click Generate to create a passphrase"
                   CaretBrush="#0078d4"/>
          <Button x:Name="btnCopy" Grid.Column="1"
                  Content="Copy"
                  Style="{StaticResource CopyBtn}"
                  VerticalAlignment="Top"
                  Margin="14,0,0,0"/>
        </Grid>
      </Border>

      <!-- == STRENGTH BAR == -->
      <Border x:Name="barTrackBorder" Grid.Row="6" CornerRadius="4" Height="8">
        <Border x:Name="strengthFill" CornerRadius="4"
                HorizontalAlignment="Left" Width="0"/>
      </Border>

      <!-- == STATS == -->
      <Grid Grid.Row="8">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="8"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="8"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="8"/>
          <ColumnDefinition Width="2*"/>
        </Grid.ColumnDefinitions>

        <Border x:Name="stat1" Grid.Column="0" BorderThickness="1" CornerRadius="8" Padding="12,10">
          <StackPanel>
            <TextBlock x:Name="lblStrengthHead" Text="STRENGTH" FontSize="9"
                       FontWeight="Bold" Margin="0,0,0,3"/>
            <TextBlock x:Name="lblStrength" Text="--" FontSize="15" FontWeight="Bold"/>
          </StackPanel>
        </Border>

        <Border x:Name="stat2" Grid.Column="2" BorderThickness="1" CornerRadius="8" Padding="12,10">
          <StackPanel>
            <TextBlock x:Name="lblEntropyHead" Text="ENTROPY" FontSize="9"
                       FontWeight="Bold" Margin="0,0,0,3"/>
            <TextBlock x:Name="lblEntropy" Text="-- bits" FontSize="15" FontWeight="Bold"/>
          </StackPanel>
        </Border>

        <Border x:Name="stat4" Grid.Column="4" BorderThickness="1" CornerRadius="8" Padding="12,10">
          <StackPanel>
            <TextBlock x:Name="lblPronounceHead" Text="PRONOUNCE" FontSize="9"
                       FontWeight="Bold" Margin="0,0,0,3"/>
            <TextBlock x:Name="lblPronounce" Text="--" FontSize="15" FontWeight="Bold"/>
          </StackPanel>
        </Border>

        <Border x:Name="stat3" Grid.Column="6" BorderThickness="1" CornerRadius="8" Padding="12,10">
          <StackPanel>
            <TextBlock x:Name="lblCrackHead" Text="TIME TO CRACK  (1T guesses/sec)" FontSize="9"
                       FontWeight="Bold" Margin="0,0,0,3"/>
            <TextBlock x:Name="lblCrack" Text="--" FontSize="15" FontWeight="Bold"/>
          </StackPanel>
        </Border>
      </Grid>

      <!-- == SETTINGS == -->
      <Border x:Name="settingsBorder" Grid.Row="10"
              BorderThickness="1" CornerRadius="10" Padding="20,16">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="28"/>
            <ColumnDefinition Width="180"/>
          </Grid.ColumnDefinitions>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="16"/>
            <RowDefinition Height="1"/>
            <RowDefinition Height="14"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <!-- Slider -->
          <StackPanel Grid.Row="0" Grid.Column="0">
            <Grid Margin="0,0,0,6">
              <TextBlock x:Name="lblWcHead" Text="WORD COUNT" FontSize="10"
                         FontWeight="Bold" VerticalAlignment="Center"/>
              <TextBlock x:Name="lblWordCount" Text="6" HorizontalAlignment="Right"
                         FontSize="22" FontWeight="Bold"/>
            </Grid>
            <Slider x:Name="sliderWords" Minimum="3" Maximum="12" Value="6"
                    IsSnapToTickEnabled="True" TickFrequency="1"/>
            <Grid Margin="2,4,2,0">
              <TextBlock x:Name="lblMin" Text="3 words" FontSize="10"/>
              <TextBlock x:Name="lblMax" Text="12 words" FontSize="10" HorizontalAlignment="Right"/>
            </Grid>
          </StackPanel>

          <!-- Separator -->
          <StackPanel Grid.Row="0" Grid.Column="2">
            <TextBlock x:Name="lblSepHead" Text="SEPARATOR" FontSize="10"
                       FontWeight="Bold" Margin="0,0,0,6"/>
            <ComboBox x:Name="cmbSeparator"/>
          </StackPanel>

          <!-- Divider -->
          <Border x:Name="divider" Grid.Row="2" Grid.ColumnSpan="3" Height="1"/>

          <!-- Options -->
          <WrapPanel Grid.Row="4" Grid.ColumnSpan="3" Orientation="Horizontal">
            <CheckBox x:Name="chkNumber" Content="Append number" Margin="0,0,28,0"/>
            <CheckBox x:Name="chkSymbol" Content="Append symbol" Margin="0,0,28,0"/>
            <CheckBox x:Name="chkAutoCopy" Content="Auto-copy on generate"/>
          </WrapPanel>
        </Grid>
      </Border>

      <!-- == GENERATE + BULK ROW == -->
      <Grid Grid.Row="12">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="8"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="8"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="8"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <Button x:Name="btnGenerate" Grid.Column="0"
                Content=">> Generate Secure Passphrase"
                Style="{StaticResource GenBtn}" Height="46"/>
        <Button x:Name="btnBulk5"   Grid.Column="2" Content="Bulk x5"
                Style="{StaticResource GhostBtn}" Height="46" Width="76"/>
        <Button x:Name="btnBulk10"  Grid.Column="4" Content="Bulk x10"
                Style="{StaticResource GhostBtn}" Height="46" Width="80"/>
        <Button x:Name="btnExport"  Grid.Column="6" Content="Export"
                Style="{StaticResource GhostBtn}" Height="46" Width="70"/>
      </Grid>

    </Grid>
  </Border>
</Window>
"@

# -- Build window --
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# -- Set window icon from passphrase.ico in same folder --
$iconPath = Join-Path $PSScriptRoot "passphrase.ico"
if (Test-Path $iconPath) {
    $window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create(
        [Uri]::new($iconPath, [UriKind]::Absolute)
    )
}

# Named elements
$rootBorder    = $window.FindName("rootBorder")
$outputBorder  = $window.FindName("outputBorder")
$barTrackBorder= $window.FindName("barTrackBorder")
$strengthFill  = $window.FindName("strengthFill")
$settingsBorder= $window.FindName("settingsBorder")
$divider       = $window.FindName("divider")
$stat1         = $window.FindName("stat1")
$stat2         = $window.FindName("stat2")
$stat3         = $window.FindName("stat3")
$stat4         = $window.FindName("stat4")

$txtPassphrase = $window.FindName("txtPassphrase")
$btnCopy       = $window.FindName("btnCopy")
$btnGenerate   = $window.FindName("btnGenerate")
$btnThemeWin   = $window.FindName("btnThemeWin")
$btnThemeKpop  = $window.FindName("btnThemeKpop")
$sliderWords   = $window.FindName("sliderWords")
$cmbSeparator  = $window.FindName("cmbSeparator")
$chkNumber     = $window.FindName("chkNumber")
$chkSymbol     = $window.FindName("chkSymbol")
$chkAutoCopy   = $window.FindName("chkAutoCopy")
$btnBulk5      = $window.FindName("btnBulk5")
$btnBulk10     = $window.FindName("btnBulk10")
$btnExport     = $window.FindName("btnExport")

$lblHeaderTop  = $window.FindName("lblHeaderTop")
$lblHeaderSub  = $window.FindName("lblHeaderSub")
$lblCharLen    = $window.FindName("lblCharLen")
$lblWordCount  = $window.FindName("lblWordCount")
$lblStrength   = $window.FindName("lblStrength")
$lblEntropy    = $window.FindName("lblEntropy")
$lblCrack      = $window.FindName("lblCrack")
$lblPronounce  = $window.FindName("lblPronounce")
$lblThemeHead  = $window.FindName("lblThemeHead")

# muted labels
$mutedLabels   = @(
    $window.FindName("lblStrengthHead"),
    $window.FindName("lblEntropyHead"),
    $window.FindName("lblCrackHead"),
    $window.FindName("lblPronounceHead"),
    $window.FindName("lblWcHead"),
    $window.FindName("lblSepHead"),
    $window.FindName("lblMin"),
    $window.FindName("lblMax"),
    $lblThemeHead
)
$checkboxes = @($chkNumber, $chkSymbol, $chkAutoCopy)

# -- Separator dropdown --
$sepMap = [ordered]@{
    "Hyphen  ( - )"     = "-"
    "Underscore  ( _ )" = "_"
    "Period  ( . )"     = "."
    "Space"             = " "
    "None"              = ""
}
foreach ($k in $sepMap.Keys) {
    $ci = New-Object System.Windows.Controls.ComboBoxItem
    $ci.Content = $k
    $ci.Tag     = $sepMap[$k]
    $cmbSeparator.Items.Add($ci) | Out-Null
}
$cmbSeparator.SelectedIndex = 0

$SYMBOLS = @('!','@','#','$','%','^','&','*','+','=','?','~')

# ==================================================================
#  APPLY THEME
# ==================================================================
function Apply-Theme {
    param([hashtable]$t)

    $window.Background         = $t.WinBg
    $rootBorder.Background     = $t.WinBg

    # output
    $outputBorder.Background   = $t.OutputBg
    $outputBorder.BorderBrush  = $t.OutputBorder
    $txtPassphrase.Foreground  = $t.OutputText
    $txtPassphrase.CaretBrush  = $t.StatValue

    # bar
    $barTrackBorder.Background = $t.BarTrack
    $strengthFill.Background   = $t.BarFill   # default; overwritten per strength

    # settings panel
    $settingsBorder.Background  = $t.PanelBg
    $settingsBorder.BorderBrush = $t.PanelBorder
    $divider.Background         = $t.DividerColor

    # stat cards
    foreach ($s in @($stat1,$stat2,$stat3,$stat4)) {
        $s.Background  = $t.PanelBg
        $s.BorderBrush = $t.PanelBorder
    }

    # header
    $lblHeaderTop.Foreground  = $t.HeaderLabel
    $lblHeaderSub.Foreground  = $t.HeaderSub
    $window.Foreground        = $t.HeaderSub

    # stat values (will be overridden by strength color, set fallback)
    $lblStrength.Foreground   = $t.StatValue
    $lblEntropy.Foreground    = $t.StatValue
    $lblCrack.Foreground      = $t.StatValue
    $lblWordCount.Foreground  = $t.CountFg
    $lblCharLen.Foreground    = $t.MutedText

    # muted text
    foreach ($lbl in $mutedLabels) {
        if ($lbl) { $lbl.Foreground = $t.MutedText }
    }

    # checkboxes
    foreach ($cb in $checkboxes) {
        $cb.Foreground = $t.MutedText
        $box = $cb.Template.FindName("box", $cb)
        $chk = $cb.Template.FindName("chk", $cb)
        if ($box) {
            $box.Background  = $t.CBBg
            $box.BorderBrush = $t.CBBorder
        }
        if ($chk) { $chk.Foreground = $t.CBCheck }
    }

    # copy button
    $btnCopy.Background  = $t.CopyBg
    $btnCopy.Foreground  = $t.CopyFg
    $btnCopy.BorderBrush = $t.CopyBorder

    # generate button
    $btnGenerate.Background  = $t.GenBg
    $btnGenerate.Foreground  = $t.GenFg
    $btnGenerate.BorderBrush = $t.GenBorder

    # ghost buttons (bulk / export)
    foreach ($gb in @($btnBulk5, $btnBulk10, $btnExport)) {
        $gb.Background  = $t.PanelBg
        $gb.Foreground  = $t.MutedText
        $gb.BorderBrush = $t.PanelBorder
    }

    # theme toggle buttons - style them as ghost
    foreach ($tb in @($btnThemeWin,$btnThemeKpop)) {
        $tb.Background  = $t.PanelBg
        $tb.Foreground  = $t.MutedText
        $tb.BorderBrush = $t.PanelBorder
    }
}

# ==================================================================
#  STATS UPDATE
# ==================================================================
function Update-Stats {
    param([string]$pp, [int]$wc)
    $entropy = Get-EntropyBits -wordCount $wc -listSize $WORDLIST.Count
    $crack   = Get-CrackTime  -entropyBits $entropy
    $label, $color, $pct = Get-StrengthInfo -entropy $entropy -t $script:activeTheme
    $pronLabel, $pronColor = Get-PronounceScore -passphrase $pp

    $trackW = $barTrackBorder.ActualWidth
    if ($trackW -le 0) { $trackW = 600 }
    $strengthFill.Width      = $trackW * $pct
    $strengthFill.Background = $color

    $lblStrength.Text        = $label
    $lblStrength.Foreground  = $color
    $lblEntropy.Text         = "$entropy bits"
    $lblEntropy.Foreground   = $color
    $lblCrack.Text           = $crack
    $lblCrack.Foreground     = $color
    $lblPronounce.Text       = $pronLabel
    $lblPronounce.Foreground = $pronColor
    $lblCharLen.Text         = "$($pp.Length) characters"
}

# ==================================================================
#  GENERATE
# ==================================================================
function Invoke-Generate {
    $sep = ($cmbSeparator.SelectedItem).Tag
    $wc  = [int]$sliderWords.Value
    $pp  = New-Passphrase -wordCount $wc -separator $sep

    if ($chkNumber.IsChecked) {
        $n  = (Get-CryptoRandom -max 9000) + 1000
        $pp = $pp + $sep + $n.ToString()
    }
    if ($chkSymbol.IsChecked) {
        $pp = $pp + $SYMBOLS[(Get-CryptoRandom -max $SYMBOLS.Count)]
    }

    $txtPassphrase.Text = $pp
    Update-Stats -pp $pp -wc $wc

    # Auto-copy if enabled
    if ($chkAutoCopy.IsChecked) {
        [System.Windows.Clipboard]::SetText($pp)
        Play-CopySound
    }
}

# ==================================================================
#  WIRE-UP EVENTS
# ==================================================================
$sliderWords.Add_ValueChanged({
    $lblWordCount.Text = [int]$sliderWords.Value
    if ($txtPassphrase.Text -notlike "Click*") { Invoke-Generate }
})

$cmbSeparator.Add_SelectionChanged({
    if ($txtPassphrase.Text -notlike "Click*") { Invoke-Generate }
})

$chkNumber.Add_Checked({   if ($txtPassphrase.Text -notlike "Click*") { Invoke-Generate } })
$chkNumber.Add_Unchecked({ if ($txtPassphrase.Text -notlike "Click*") { Invoke-Generate } })
$chkSymbol.Add_Checked({   if ($txtPassphrase.Text -notlike "Click*") { Invoke-Generate } })
$chkSymbol.Add_Unchecked({ if ($txtPassphrase.Text -notlike "Click*") { Invoke-Generate } })

$btnGenerate.Add_Click({ Invoke-Generate })

$window.Add_KeyDown({
    if ($_.Key -eq [System.Windows.Input.Key]::Return) { Invoke-Generate }
})

$window.Add_SizeChanged({
    if ($txtPassphrase.Text -notlike "Click*") {
        Update-Stats -pp $txtPassphrase.Text -wc ([int]$sliderWords.Value)
    }
})

# Copy button
$btnCopy.Add_Click({
    $text = $txtPassphrase.Text
    if ($text -and $text -notlike "Click*") {
        [System.Windows.Clipboard]::SetText($text)
        Play-CopySound
        $btnCopy.Content    = "Copied!"
        $btnCopy.Foreground = "#4ade80"
        $null = $window.Dispatcher.BeginInvoke([Action]{
            Start-Sleep -Milliseconds 1800
            $window.Dispatcher.Invoke([Action]{
                $btnCopy.Content    = "Copy"
                $btnCopy.Foreground = $script:activeTheme.CopyFg
            })
        }, [System.Windows.Threading.DispatcherPriority]::Background)
    }
})

# ==================================================================
#  BULK GENERATE  - opens a picker window
# ==================================================================
function Show-BulkWindow {
    param([int]$count)

    $sep = ($cmbSeparator.SelectedItem).Tag
    $wc  = [int]$sliderWords.Value
    $t   = $script:activeTheme

    # Generate the batch
    $batch = for ($i = 0; $i -lt $count; $i++) {
        $pp = New-Passphrase -wordCount $wc -separator $sep
        if ($chkNumber.IsChecked) {
            $n  = (Get-CryptoRandom -max 9000) + 1000
            $pp = $pp + $sep + $n.ToString()
        }
        if ($chkSymbol.IsChecked) {
            $pp = $pp + $SYMBOLS[(Get-CryptoRandom -max $SYMBOLS.Count)]
        }
        $pp
    }

    # Build bulk window
    $bw = New-Object System.Windows.Window
    $bw.Title               = "Bulk Generate - Pick One"
    $bw.Width               = 660
    $bw.Height              = 480
    $bw.MinWidth            = 480
    $bw.MinHeight           = 340
    $bw.Background          = $t.WinBg
    $bw.WindowStartupLocation = "CenterOwner"
    $bw.Owner               = $window
    $bw.ResizeMode          = "CanResizeWithGrip"

    $iconPath = Join-Path $PSScriptRoot "passphrase.ico"
    if (Test-Path $iconPath) {
        $bw.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create(
            [Uri]::new($iconPath, [UriKind]::Absolute))
    }

    $outerGrid = New-Object System.Windows.Controls.Grid
    $outerGrid.Margin = "20,16,20,16"

    $r0 = New-Object System.Windows.Controls.RowDefinition; $r0.Height = "Auto"
    $r1 = New-Object System.Windows.Controls.RowDefinition; $r1.Height = "8"
    $r2 = New-Object System.Windows.Controls.RowDefinition; $r2.Height = "*"
    $r3 = New-Object System.Windows.Controls.RowDefinition; $r3.Height = "10"
    $r4 = New-Object System.Windows.Controls.RowDefinition; $r4.Height = "Auto"
    $outerGrid.RowDefinitions.Add($r0)
    $outerGrid.RowDefinitions.Add($r1)
    $outerGrid.RowDefinitions.Add($r2)
    $outerGrid.RowDefinitions.Add($r3)
    $outerGrid.RowDefinitions.Add($r4)

    $hdr = New-Object System.Windows.Controls.TextBlock
    $hdr.Text       = "Click any passphrase to select it and copy to clipboard"
    $hdr.FontSize   = 12
    $hdr.Foreground = $t.MutedText
    [System.Windows.Controls.Grid]::SetRow($hdr, 0)
    $outerGrid.Children.Add($hdr) | Out-Null

    $listBorder = New-Object System.Windows.Controls.Border
    $listBorder.Background   = $t.PanelBg
    $listBorder.BorderBrush  = $t.PanelBorder
    $listBorder.BorderThickness = "1"
    $listBorder.CornerRadius = "8"
    [System.Windows.Controls.Grid]::SetRow($listBorder, 2)

    $sv = New-Object System.Windows.Controls.ScrollViewer
    $sv.VerticalScrollBarVisibility = "Auto"
    $sv.Margin = "4"

    $sp = New-Object System.Windows.Controls.StackPanel
    $sp.Margin = "8,4,8,4"

    $selectedPP = $null

    foreach ($pp in $batch) {
        $pron, $pronCol = Get-PronounceScore -passphrase $pp

        $rowBorder = New-Object System.Windows.Controls.Border
        $rowBorder.CornerRadius    = "6"
        $rowBorder.Padding         = "12,10"
        $rowBorder.Margin          = "0,3,0,3"
        $rowBorder.Background      = $t.OutputBg
        $rowBorder.BorderBrush     = $t.OutputBorder
        $rowBorder.BorderThickness = "1"
        $rowBorder.Cursor          = "Hand"
        $rowBorder.Tag             = $pp

        $rowGrid = New-Object System.Windows.Controls.Grid
        $c0 = New-Object System.Windows.Controls.ColumnDefinition; $c0.Width = "*"
        $c1 = New-Object System.Windows.Controls.ColumnDefinition; $c1.Width = "Auto"
        $rowGrid.ColumnDefinitions.Add($c0)
        $rowGrid.ColumnDefinitions.Add($c1)

        $ppTxt = New-Object System.Windows.Controls.TextBlock
        $ppTxt.Text       = $pp
        $ppTxt.FontFamily = "Cascadia Mono,Consolas,monospace"
        $ppTxt.FontSize   = 14
        $ppTxt.Foreground = $t.OutputText
        $ppTxt.TextWrapping = "Wrap"
        [System.Windows.Controls.Grid]::SetColumn($ppTxt, 0)

        $pronTxt = New-Object System.Windows.Controls.TextBlock
        $pronTxt.Text       = $pron
        $pronTxt.FontSize   = 11
        $pronTxt.FontWeight = "Bold"
        $pronTxt.Foreground = $pronCol
        $pronTxt.VerticalAlignment = "Center"
        $pronTxt.Margin    = "12,0,0,0"
        [System.Windows.Controls.Grid]::SetColumn($pronTxt, 1)

        $rowGrid.Children.Add($ppTxt)  | Out-Null
        $rowGrid.Children.Add($pronTxt)| Out-Null
        $rowBorder.Child = $rowGrid

        # Hover effects
        $rowBorder.Add_MouseEnter({
            $this.Background = $script:activeTheme.PanelBorder
        })
        $rowBorder.Add_MouseLeave({
            $this.Background = $script:activeTheme.OutputBg
        })

        # Click: use this passphrase
        $rowBorder.Add_MouseLeftButtonUp({
            $chosen = $this.Tag
            $txtPassphrase.Text = $chosen
            Update-Stats -pp $chosen -wc ([int]$sliderWords.Value)
            [System.Windows.Clipboard]::SetText($chosen)
            Play-CopySound
            $bw.Close()
        })

        $sp.Children.Add($rowBorder) | Out-Null
    }

    $sv.Content        = $sp
    $listBorder.Child  = $sv
    $outerGrid.Children.Add($listBorder) | Out-Null

    # Close button
    $closeBtn = New-Object System.Windows.Controls.Button
    $closeBtn.Content         = "Cancel"
    $closeBtn.Height          = 36
    $closeBtn.FontSize        = 12
    $closeBtn.Background      = $t.PanelBg
    $closeBtn.Foreground      = $t.MutedText
    $closeBtn.BorderBrush     = $t.PanelBorder
    $closeBtn.BorderThickness = "1"
    $closeBtn.Cursor          = "Hand"
    $closeBtn.Add_Click({ $bw.Close() })
    [System.Windows.Controls.Grid]::SetRow($closeBtn, 4)
    $outerGrid.Children.Add($closeBtn) | Out-Null

    $bw.Content = $outerGrid
    $bw.ShowDialog() | Out-Null
}

$btnBulk5.Add_Click({  Show-BulkWindow -count 5  })
$btnBulk10.Add_Click({ Show-BulkWindow -count 10 })

# ==================================================================
#  EXPORT  - save a batch to a .txt file
# ==================================================================
$btnExport.Add_Click({
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Title      = "Export Passphrases"
    $dlg.Filter     = "Text files (*.txt)|*.txt"
    $dlg.FileName   = "passphrases_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

    if ($dlg.ShowDialog() -eq $true) {
        $sep = ($cmbSeparator.SelectedItem).Tag
        $wc  = [int]$sliderWords.Value
        $lines = @("Passphrase Export - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        $lines += "Word count: $wc  |  Separator: '$sep'  |  Wordlist: $($WORDLIST.Count) words"
        $lines += "=" * 60
        $lines += ""

        for ($i = 1; $i -le 20; $i++) {
            $pp = New-Passphrase -wordCount $wc -separator $sep
            if ($chkNumber.IsChecked) {
                $n  = (Get-CryptoRandom -max 9000) + 1000
                $pp = $pp + $sep + $n.ToString()
            }
            if ($chkSymbol.IsChecked) {
                $pp = $pp + $SYMBOLS[(Get-CryptoRandom -max $SYMBOLS.Count)]
            }
            $pron, $_ = Get-PronounceScore -passphrase $pp
            $lines += "$("{0:D2}" -f $i).  $pp   [$pron]"
        }

        $lines += ""
        $lines += "Generated by Passphrase Generator"
        $lines | Out-File -FilePath $dlg.FileName -Encoding UTF8

        # Confirm
        $btnExport.Content    = "Exported!"
        $btnExport.Foreground = "#4ade80"
        $null = $window.Dispatcher.BeginInvoke([Action]{
            Start-Sleep -Milliseconds 2000
            $window.Dispatcher.Invoke([Action]{
                $btnExport.Content    = "Export"
                $btnExport.Foreground = $script:activeTheme.MutedText
            })
        }, [System.Windows.Threading.DispatcherPriority]::Background)
    }
})

# Theme buttons
$btnThemeWin.Add_Click({
    $script:activeTheme = $THEMES[(Get-WindowsTheme)]
    Apply-Theme $script:activeTheme
    if ($txtPassphrase.Text -notlike "Click*") {
        Update-Stats -pp $txtPassphrase.Text -wc ([int]$sliderWords.Value)
    }
})

$btnThemeKpop.Add_Click({
    $script:activeTheme = $THEMES["KPopDemon"]
    Apply-Theme $script:activeTheme
    if ($txtPassphrase.Text -notlike "Click*") {
        Update-Stats -pp $txtPassphrase.Text -wc ([int]$sliderWords.Value)
    }
})

# -- On load: detect Windows theme, apply, generate --
$window.Add_Loaded({
    $script:activeTheme = $THEMES[(Get-WindowsTheme)]
    Apply-Theme $script:activeTheme
    Invoke-Generate
})

$window.ShowDialog() | Out-Null
