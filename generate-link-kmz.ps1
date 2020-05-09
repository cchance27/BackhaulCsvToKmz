$sites = Import-Csv -Delimiter ',' -Path sites.csv
$links = Import-Csv -Delimiter ',' -Path links.csv

$output = Join-Path -Path (Get-Location) -ChildPath "$(get-date -Format "yyyy-MM-dd") - Backhauls.kml"

Write-Host "Backup Old KML"
Move-Item -Path $output -Destination $output".backup" -force -ErrorAction SilentlyContinue | Out-Null

$lineWidth120 = 8
$lineWidth40 = 2
$channelColors = @{
	'10735'='da55c1';	
	'10775'='cef10e';	
	'10815'='2e22fd';	
	'10855'='c2bc1b';	
	'10895'='72e3a9';	
	'10935'='208cbc';	
	'10975'='de84cf';	
	'11015'='cf079f';	
	'11055'='ffada5';	
	'11095'='90f3b9';	
	'11135'='a567f5';	
	'11175'='2cddf6';
} 

$kml = @"
<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2" xmlns:gx="http://www.google.com/kml/ext/2.2" xmlns:kml="http://www.opengis.net/kml/2.2" xmlns:atom="http://www.w3.org/2005/Atom">
    <Document>
    	<Style id="site-visible">
		    <IconStyle>
		    	<color>ffffaa00</color>
		    	<Icon>
		    		<href>http://maps.google.com/mapfiles/kml/paddle/blu-blank.png</href>
		    	</Icon>
		    	<hotSpot x="32" y="1" xunits="pixels" yunits="pixels"/>
		    </IconStyle>
		    <LabelStyle>
		    	<color>ffffaa00</color>
		    	<scale>1.0</scale>
		    </LabelStyle>
		    <LineStyle>
		    	<color>ffffaa00</color>
		    </LineStyle>
	    </Style>
	    <Style id="site-hidden">
	    	<IconStyle>
	    		<color>ffffaa00</color>
	    		<Icon>
	    			<href>http://maps.google.com/mapfiles/kml/paddle/blu-blank.png</href>
	    		</Icon>
	    		<hotSpot x="32" y="1" xunits="pixels" yunits="pixels"/>
	    	</IconStyle>
	    	<LabelStyle>
	    		<color>00ffaa00</color>
	    		<scale>1.0</scale>
	    	</LabelStyle>
	    	<LineStyle>
	    		<color>00ffaa00</color>
	    	</LineStyle>
	    </Style>
	    <StyleMap id="site">
	    	<Pair>
	    		<key>normal</key>
	    		<styleUrl>#site-hidden</styleUrl>
	    	</Pair>
	    	<Pair>
	    		<key>highlight</key>
	    		<styleUrl>#site-visible</styleUrl>
	    	</Pair>
	    </StyleMap>
        <name>UTS-11GHz</name>
        <Folder>
            <name>Sites</name>
            $(
            	$sites | foreach-object {
					Write-Host "New Site:" $_.Site

            		'<Placemark>
            		   <name>{0}</name>
            		   <styleUrl>#site</styleUrl>
            		   <Point>
            		   	<extrude>1</extrude>
	    			       <altitudeMode>relativeToGround</altitudeMode>
            		       <coordinates>{1},{2},0</coordinates>
            		   </Point>
            		</Placemark>
            		' -f "[$($_.BBCode)] $($_.Site)", $_.Latitude, $_.Longitude
             	}
            )
        </Folder>
        <Folder>
            <name>Backhauls</name>
            $(
				$specialFields = ($links | get-member).Name | Where-Object { $_.ToString().StartsWith("_") }
				
                $links | foreach-object {
					Write-Host "New Link: "$fromSite.Site"-"$toSite.Site

                    $link = $_
                    $fromSite = $($sites | Where-Object { $_.BBCode -eq $link.From })
					$toSite = $($sites | Where-Object { $_.BBCode -eq $link.To })
					if ($link."Channel Size" -eq 120) { $linkWidth = $lineWidth120 } else { $linkWidth = $lineWidth40 }

					$lowBandChannel = if ($link."_Channel TX" -lt $link."_Channel RX") { $link."_Channel TX" } else { $link."_Channel RX" }
					$linkColor = $channelColors[$lowbandChannel]

					$specialData = '<table border=0 width=100%>'
					$specialData += '<tr><td><b>From</b></td><td>{0}</td></tr>' -f $fromSite.Site
					$specialData += '<tr><td><b>To</b></td><td>{0}</td></tr>' -f $toSite.Site
					$specialData += $($specialFields | foreach-object { 
						'<tr><td><b>{0}</b></td><td>{1}</td></tr>' -f $_.replace("_", ""), $link."$($_)"
					} | Out-String)
					$specialData += '</table>'
					
                    '<Placemark> 
			            <name>{0}-{1}</name> 
			            <description><![CDATA[{7}]]></description>
			            <LineString>
			        	    <coordinates>
			        			{2},{3},0.
                            	{4},{5},0.
       		        		</coordinates>
			            	<altitudeMode>relativeToGround</altitudeMode>
							<extrude>1</extrude>
						</LineString>
						<Style>
	    					<LineStyle>  
	    						<color>#ff{11}</color>
	    						<width>{10}</width>
	    					</LineStyle> 
	    				</Style>
		             </Placemark>
                    ' -f $link.From, 
                         $link.To, 
                         $fromSite.Latitude, 
                         $fromSite.Longitude, 
                         $toSite.Latitude, 
                         $toSite.Longitude,
                         $link."Channel Size",
                         $specialData,
                         $fromSite.Site,
						 $toSite.Site,
						 $linkWidth,
						 $linkColor
                  })
        </Folder>
    </Document>
</kml>
"@

$kml | Out-File -Encoding utf8 ($output)
