# 소스 폴더와 대상 폴더를 정의합니다.
$bksources = @(
    "\\geomsa-main2\hfa",
    "\\orb\orb",
    "\\192.168.0.221\pdf",
    "\\topo\Topo",
    "\\topo\TMS_Backup",
    "\\topo\tms2021",
    "\\AFC-210-PC\Fundus2",
    "\\oqas\oqass",
    "\\antoct\antoct",
    "\\slitlamp\slitlamp",
    "C:\sp"
)

$destinations = @(
    "D:\BACKUP\hfa",
    "D:\BACKUP\ORB",
    "D:\BACKUP\OCT\2022.06~",
    "D:\BACKUP\TOPO",
    "D:\BACKUP\TOPO Tms_Backup",
    "D:\BACKUP\TOPO Tms\2021",
    "D:\BACKUP\FUNDUS\2017.7~",
    "D:\BACKUP\Oqas",
    "D:\BACKUP\Antoct",
    "D:\BACKUP\Slitlamp",
    "D:\BACKUP\SP"
)

# 어제 생성된 파일을 복사합니다.
$yesterday = (Get-Date).AddDays(-1)

# 작업 완료 여부를 나타내는 변수를 초기화합니다.
$taskCompleted = $true

# 로그 파일을 저장할 폴더를 지정합니다.
$logFolder = "D:\BACKUP"

foreach ($source in $bksources) {
    $sourceFolder = $source
    $destinationFolder = $destinations[$bksources.IndexOf($source)]

    # 로그 파일 경로를 설정합니다.
    $logFileName = Join-Path $logFolder "$((Split-Path $destinationFolder -Leaf).ToLower()).txt"

    # 소스 폴더에서 어제 생성된 파일을 찾아 대상 폴더로 복사합니다.
    $filesToCopy = Get-ChildItem $sourceFolder -Recurse | Where-Object { $_.CreationTime.Date -eq $yesterday.Date -and -not $_.PSIsContainer }
    
    # 작업 진행 상황을 윈도우 창에 표시합니다.
    Write-Progress -Activity "Copying files from $sourceFolder" -Status "In progress" -PercentComplete 0
    $totalCount = $filesToCopy.Count
    $index = 0

    foreach ($file in $filesToCopy) {
        $destinationFile = Join-Path $destinationFolder $file.Name
        # 복사 작업
        try {
            Copy-Item $file.FullName $destinationFile -Force
            Write-Host "파일 복사 완료: $($file.FullName)"
            # 성공 정보를 로그 파일에 기록합니다.
            Add-Content -Path $logFileName -Value ("Success: $($file.FullName)")
        }
        catch {
            Write-Host "파일 복사 실패: $($file.FullName)"
            $taskCompleted = $false  # 작업이 실패하면 $taskCompleted를 false로 설정합니다.
            # 실패 정보를 로그 파일에 기록합니다.
            Add-Content -Path $logFileName -Value ("Failure: $($file.FullName)")
        }

        # 작업 진행 상황을 업데이트합니다.
        $index++
        $percentComplete = [math]::Round(($index / $totalCount) * 100)
        Write-Progress -Activity "Copying files from $sourceFolder" -Status "In progress" -PercentComplete $percentComplete
    }
}

# 작업 완료 메시지를 표시합니다.
if ($taskCompleted) {
    $message = "작업이 완료되었습니다."
    $title = "작업 완료"
    Show-Message -message $message -title $title -success $true
} else {
    $message = "작업이 실패했습니다."
    $title = "작업 실패"
    Show-Message -message $message -title $title -success $false
}
