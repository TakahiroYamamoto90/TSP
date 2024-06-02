Param(
    [Parameter()]
    [String]$INPUT_CSV_FILE_PATH = "D:\PATH\TO\課題3-3.csv",
    [String]$OUTPUT_CSV_FOLDER_PATH = "D\PATH\TO\OUTPUT_FOLDER"
)

# XY座標の2点間の距離を求める
function CalculateDistance {
    param(
        [double]$x1,
        [double]$y1,
        [double]$x2,
        [double]$y2
    )

    $dx = $x2 - $x1
    $dy = $y2 - $y1
    return [Math]::Sqrt($dx * $dx + $dy * $dy)
}

# 次のルートを返す
function GetNextRoute {
    param(
        [pscustomobject]$csv_row,
        [int]$offset = 0,
        [string]$strategy = "greedy",
        [System.Array]$base_route_towns = @()
    )

    if ($strategy -eq "greedy") {
        return GetGreedyRoute -csv_row $csv_row -offset $offset
    } elseif($strategy -eq "hill_climb") {
        return GetHillClimbRoute -csv_row $csv_row -base_route_towns $base_route_towns -offset $offset
    }
    
}

# 既存のルートを受けて、次の最短距離のルートを返す
function GetHillClimbRoute {
    param(
        [pscustomobject]$csv_row,
        [int]$offset = 0,
        [System.Array]$base_route_towns = @()
    )

    $min_distance = 999
    $min_town = ""
    $max_distance = 0
    $max_town = ""

    $town_distances = @()
    # ベースルートから、まだ通っていない次のルートを列挙する
    foreach($base_route_town in $base_route_towns) {
        if ($base_route_town.town -eq  $csv_row.H1) {
            continue
        }
        if ($towns_kvs[$base_route_town.town].once_passed -eq $true) {
            continue
        }
        $town_distance = [PSCustomObject]@{
            town = $base_route_town.town;
            distance = $csv_row.($base_route_town.town);
        }
        $town_distances += $town_distance
    }

    $town_distances = $town_distances | Sort-Object -Property "distance"
    $min_town = $town_distances[$offset]
    $max_town = $town_distances[$town_distances.Length - 1 - $offset]

    if ([string]::IsNullOrEmpty($town_distances[$offset]) -eq $true) {
        $min_town = $town_distances[0]
    }

    if ([string]::IsNullOrEmpty($town_distances[$town_distances.Length - 1 - $offset]) -eq $true) {
        $max_town = $town_distances[0]
    }

    $towns_kvs[$csv_row.H1].once_passed = $true

    return [pscustomobject]@{
        min_town=$min_town.town;
        min_distance=$min_town.distance;
        max_town=$max_town.town;
        max_distance=$max_town.distance;
    }

}

# 欲張り法での、次の最短距離のルートを返す
function GetGreedyRoute {
    param(
        [pscustomobject]$csv_row,
        [int]$offset = 0
    )

    $min_distance = 999
    $min_town = ""
    $max_distance = 0
    $max_town = ""

    $town_distances = @()
    foreach($town in $towns) {
        if ($town -eq  $csv_row.H1) {
            continue
        }
        if ($towns_kvs[$town].once_passed -eq $true) {
            continue
        }
        $town_distance = [PSCustomObject]@{
            town = $town;
            distance = $csv_row."$town";
        }
        $town_distances += $town_distance
    }

    $town_distances = $town_distances | Sort-Object -Property "distance"
    $min_town = $town_distances[$offset]
    $max_town = $town_distances[$town_distances.Length - 1 - $offset]

    if ([string]::IsNullOrEmpty($town_distances[$offset]) -eq $true) {
        $min_town = $town_distances[0]
    }

    if ([string]::IsNullOrEmpty($town_distances[$town_distances.Length - 1 - $offset]) -eq $true) {
        $max_town = $town_distances[0]
    }

    $towns_kvs[$csv_row.H1].once_passed = $true

    return [pscustomobject]@{
        min_town=$min_town.town;
        min_distance=$min_town.distance;
        max_town=$max_town.town;
        max_distance=$max_town.distance;
    }
}

# ルート探索を実行する
function RouteSearch {
    param(
        [System.Array]$hc_target_towns = @(),
        [string]$strategy = "greedy",
        [System.Array]$base_route_towns = @()
    )

    $csv_rows = import-csv -path $INPUT_CSV_FILE_PATH -Encoding UTF8
    $towns = @()
    $towns_kvs = @{}

    foreach($csv_row1 in $csv_rows) {
        foreach($csv_row2 in $csv_rows) {
            $distance = "999"
            if ($csv_row1.H1 -ne $csv_row2.H1) {
                $distance = CalculateDistance -x1 $csv_row1.X -y1 $csv_row1.Y -x2 $csv_row2.X  -y2 $csv_row2.Y
            }
            Add-Member -InputObject $csv_row1 -MemberType NoteProperty -Name $csv_row2.H1 -Value $distance
        }
        $towns += $csv_row1.H1
        $towns_kvs[$csv_row1.H1] = [pscustomobject]@{
            csv_row=$csv_row1;
            once_passed=$false;
        }
    }

    $csv_rows | export-csv -Path "$OUTPUT_CSV_FOLDER_PATH\課題3-3_距離追加.csv" -Encoding UTF8 -NoTypeInformation

    $result = $null
    $distance_sum = 0
    $passed_towns = @()
    $town_sum = @("王都")
    $passed_count = 0

    $towns_kvs["王都"].once_passed = $true
    $result = [pscustomobject]@{
        min_town="王都";
        min_distance=0;
        max_town="王都";
        max_distance=0;
    }

    while ($towns_kvs.GetEnumerator().where({$_.Value.once_passed -eq $false}).count -gt 1) {
        $next_town = $towns_kvs[$result.min_town].csv_row
        $result = GetNextRoute -csv_row $next_town -offset ([int]$hc_target_towns.Contains($result.min_town)) -strategy $strategy -base_route_towns $base_route_towns
        $distance_sum += $result.min_distance
        $town_sum += $result.min_town
        $passed_count++
        $passed_towns += [pscustomobject]@{town=$result.min_town;distance=$result.min_distance;passed_count=$passed_count;distance_sum=$distance_sum;town_sum=($town_sum -join "->");}
    }

    $towns_kvs["王都"].once_passed = $false
    $next_town = $towns_kvs[$result.min_town].csv_row
    $result = GetNextRoute -csv_row $next_town -offset ([int]$hc_target_towns.Contains($result.min_town)) -strategy $strategy -base_route_towns $base_route_towns
    $distance_sum += $result.min_distance
    $town_sum += $result.min_town
    $passed_count++
    $passed_towns += [pscustomobject]@{town=$result.min_town;distance=$result.min_distance;passed_count=$passed_count;distance_sum=$distance_sum;town_sum=($town_sum -join "->");}

    return [pscustomobject]@{distance_sum=$distance_sum;passed_towns=$passed_towns;}
}

### ここからmain処理開始 ###

$now_string = get-date -Format "yyyyMMdd hhmmss"
$main_csv_rows = import-csv -path $INPUT_CSV_FILE_PATH -Encoding UTF8
$main_towns = @()
foreach($main_csv_row in $main_csv_rows) {
    $main_towns += $main_csv_row.H1
}

$all_results = @()
$best_results = @()
$greedy_result = RouteSearch -hc_target_towns @()
$all_results += [pscustomobject]@{
    try_cnt=0;
    hc_target_towns_cnt=0;
    hc_target_towns=(@() -join ",");
    two_opt_target_towns="";
    distance_summary=$greedy_result.distance_sum;
    baseline_diff=0;
    best_diff=0;
    town_sum=$greedy_result.passed_towns[19].town_sum;
}

$best_results += [pscustomobject]@{
    try_cnt=0;
    hc_target_towns_cnt=0;
    hc_target_towns=(@() -join ",");
    two_opt_target_towns="";
    distance_summary=$greedy_result.distance_sum;
    baseline_diff=0;
    best_diff=0;
    town_sum=$greedy_result.passed_towns[19].town_sum;
}

Format-Table -InputObject $greedy_result.passed_towns

$best_result = $greedy_result.psobject.Copy()
$hc_target_towns_strings = @()

for($i = 1;$i -le 1000;$i++) {
    do {
        $hc_target_towns = Get-Random $main_towns[0..14] -Count (Get-Random -Minimum 1 -Maximum 15) | Sort-Object
        $hc_target_towns_string = $hc_target_towns -join ","
    } while ($hc_target_towns_strings.Contains($hc_target_towns_string) -eq $true)
    $hc_target_towns_strings += $hc_target_towns_string

    "hc_target_towns:" + ($hc_target_towns -join ",")
    $hc_result = RouteSearch -hc_target_towns $hc_target_towns -strategy "hill_climb" -base_route_towns $best_result.passed_towns

    "最新結果距離:" + $hc_result.distance_sum.ToString() + " / " + "ベスト結果距離:" + $best_result.distance_sum.ToString()
    Format-Table -InputObject $hc_result.passed_towns

    if ($hc_result.distance_sum -lt $best_result.distance_sum) {
        $best_results += [pscustomobject]@{
            try_cnt=$i;
            hc_target_towns_cnt=$hc_target_towns.Count;
            hc_target_towns=($hc_target_towns -join ",");
            two_opt_target_towns="";
            distance_summary=$hc_result.distance_sum;
            baseline_diff=($hc_result.distance_sum-$greedy_result.distance_sum);
            best_diff=($hc_result.distance_sum-$best_result.distance_sum);
            town_sum=$hc_result.passed_towns[19].town_sum;
        }
        $best_result = $hc_result
    }

    $all_results += [pscustomobject]@{
        try_cnt=$i;
        hc_target_towns_cnt=$hc_target_towns.Count;
        hc_target_towns=($hc_target_towns -join ",");
        two_opt_target_towns="";
        distance_summary=$hc_result.distance_sum;
        baseline_diff=($hc_result.distance_sum-$greedy_result.distance_sum);
        best_diff=($hc_result.distance_sum-$best_result.distance_sum);
        town_sum=$hc_result.passed_towns[19].town_sum;
    }
}

Format-Table -InputObject ($all_results | Sort-Object -Property "hc_target_towns" | Get-Unique -AsString | Sort-Object -Property "distance_summary")
$all_results | export-csv -Path "$OUTPUT_CSV_FOLDER_PATH \課題3-3_hillclimb_all_$now_string.csv" -Encoding UTF8 -NoTypeInformation

Format-Table -InputObject ($best_results | Sort-Object -Property "hc_target_towns" | Get-Unique -AsString | Sort-Object -Property "distance_summary")
$best_results | export-csv -Path "$OUTPUT_CSV_FOLDER_PATH \課題3-3_hillclimb_best_$now_string.csv" -Encoding UTF8 -NoTypeInformation