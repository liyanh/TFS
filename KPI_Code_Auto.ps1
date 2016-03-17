$resp = curl -usedefaultCredentials https://tfs.slb.com/tfs/SLB1/Prism/_apis/git/repositories?api-version=1.0
$li = ($resp.Content|ConvertFrom-Json).value.name
$exclude = "Archive","Core.Service.DemoApp","TechnicalSpike","WIRT"
foreach($na in $li){
	if($na -in $exclude){
		$na
		continue
	}
	if(!($na.StartsWith("Poc.") -or $na.StartsWith("RO.") -or $na.StartsWith("WI.") -or $na.StartsWith("Rhapsody."))){
		$path = "\\ia-env\PrismMetrics\src\"+$na
		if(![io.Directory]::Exists($path)){
			$addr = "https://tfs.slb.com/tfs/SLB1/Prism/_git/"+$na
			git clone $addr
		}else{
			pushd $na
			git pull --reb
			popd
		}
	}
}
pause