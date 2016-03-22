$resp = curl -usedefaultCredentials https://tfs.slb.com/tfs/SLB1/Prism/_apis/git/repositories?api-version=1.0
$li = ($resp.Content|ConvertFrom-Json).value.name
$exclude = "Archive","Core.Service.DemoApp","TechnicalSpike","WIRT"
$excludeLike = "Poc.","RO.","WI.","Rhapsody."
[System.Collections.ArrayList]$res = $li
foreach($na in $li){
	if($na -in $exclude){
		$res.Remove($na)
		continue
	}
	foreach($exli in $excludeLike){
		if($na.StartsWith($exli)){
			$res.Remove($na)
		}
	}	
}
[System.Collections.ArrayList]$exis = Get-ChildItem
foreach($fil in $exis){
	if(!($fil.name -in $res) -and ($fil -is [IO.DirectoryInfo])){
		Remove-Item -Recurse -Force $fil
	}
}
foreach($fet in $res){
	$path = ".\"+$fet
	if(![io.Directory]::Exists($path)){
		$addr = "https://tfs.slb.com/tfs/SLB1/Prism/_git/"+$fet
		git clone $addr
	}else{
		pushd $fet
		git pull --reb
		popd
	}
}
