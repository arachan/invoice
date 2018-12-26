# load data
$data=Import-Csv -Path .\data.csv -Encoding UTF8

# load template
$temp=Get-Content .\invoice.fodg -Encoding UTF8

# cnt row in report
$i=1
# page
$p=0
# reportrows
$rprow=6
# total row count
$rc=0

# calcurate total
$sum=$data | Measure-Object 'subtotal' -sum

foreach($d in $data){     

    # merge template
    $temp=$temp -replace '{salesno}',$d.salesno
    $temp=$temp -replace '{customer}',$d.customer
    $temp=$temp -replace '{supplier}',$d.supplier
    $temp=$temp -replace "{item$i}",$d.item
    $temp=$temp -replace "{subtotal$i}",$d.subtotal
    
    # rowcount
    $rc=$rc+1

    # report end write build report
    if($i%$rprow -eq 0){
        # Write-Output "行"$rc
        # Write-Output "data数"$data.Count
        
        # Page Count
        $p=$p+1

        # write total
        if($rc -eq $data.Count){
            $temp=$temp -replace '{total}',$sum.Sum
        }

        # add Page No
        $temp=$temp -replace '{page}',"-$p"
        # delete {}
        $temp=$temp -replace "{(.*)}",""

        # output name 
        $date=Get-Date -Format yyyyMMddhhmmss
        $file="output\"+$date+$p

        # output newfile
        $temp | out-file -FilePath $file'.fodg' -Encoding utf8
        echo "1ページ目"
        $i=$i-$rprow
        echo $i
        # load template
        $temp=Get-Content .\invoice.fodg -Encoding UTF8
    }
    
    # report in row
    $i=$i+1
    echo $i
    #echo $rc
    
}

# output remainder data to report
if($data.Count%$rprow -ne 0){
    $p=$p+1
    $file="output\"+$date+$p

    # write total
    $temp=$temp -replace '{total}',$sum.Sum
    $temp=$temp -replace '{page}',"-$p"
    
    # delete {}
    $temp=$temp -replace "{(.*)}","" 
    
    #output report
    $temp | out-file -FilePath $file'.fodg' -Encoding utf8

}