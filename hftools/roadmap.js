var appData = {files:{}};

function toggleOverlay(toggle, msg){
	document.getElementById("uiblock").style.display = toggle;
	$("#uiblock").html(msg)
}

function openTab(evt, tabId) {
	$(".tabs").hide();
	$(".tablink").removeClass("w3-red");
	$("#"+tabId).show();
	if(evt) evt.currentTarget.className += " w3-red";
}

function uploadFile(type) {
	if(appData[type] && appData[type].length > 1){
		if(type == "F") alert("Features already uploaded!");
		if(type == "B") alert("Backlogs already uploaded!");
		return;
	}
	appData.importType = type;
	$('#importFile').click();
}

function processImport(){
	var file = $('#importFile')[0].files[0];
	var reader = new FileReader();
	reader.fileName = file.name;
	reader.onload = function(e){
		var workbook = XLSX.read(e.target.result);
		workbook.SheetNames.forEach(function(sheet){
			var xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet], {raw:false, header:1});
			if((appData.importType == "F" && xlData.length > 1 && xlData[0].length > 1 && xlData[0][0].toUpperCase() == "EPIC ID") || (appData.importType == "B" && xlData.length > 1 && xlData[0].length > 1 && xlData[0][0].toUpperCase() == "ID")){
				appData[appData.importType] = xlData;				
				appData.importType = null;				
			} else {
				if(appData.importType == "F") alert("Please upload a valid Features file.");
				if(appData.importType == "B") alert("Please upload a valid Backlogs file.");
				return;
			}
			processData();
		})
	}
	reader.readAsArrayBuffer(file, "UTF-8");
	$('#importFile').val("");
}

function processData(){
	if(!appData['F']){
		alert("Please upload Features to generate Roadmap");
		return;
	}
	if(!appData['B']){
		alert("Please upload Backlogs to generate Roadmap");
		return;
	}
	if(appData['F'] && appData['B']){
		toggleOverlay('block');
		initFeaturesTable();
		initBacklogsTable();
		genRoadmapData();
		toggleOverlay('none');
	}
	console.log(appData);
}

function genRoadmapData(){
	appData.roadMap = {};
	for(var i = 1; i<appData.F.length; i++){
		if(!appData.roadMap[appData.F[i][6]]) appData.roadMap[appData.F[i][6]] = {epics:{}, itr:{}};
		if(!appData.roadMap[appData.F[i][6]].epics[appData.F[i][0]]){
			appData.roadMap[appData.F[i][6]].epics[appData.F[i][0]] = {name:appData.F[i][1], features:{}, f_count:0, s_count:0, s_points:0}
		}
		if(!appData.roadMap[appData.F[i][6]].epics[appData.F[i][0]].features[appData.F[i][3]]){
			appData.roadMap[appData.F[i][6]].epics[appData.F[i][0]].f_count++;
			appData.roadMap[appData.F[i][6]].epics[appData.F[i][0]].features[appData.F[i][3]] = {name:appData.F[i][2], dependency:appData.F[i][5] ? appData.F[i][5] : "", team:appData.F[i][8], status:appData.F[i][18], itr:{}, s_count:0, s_points:0}
		}
	}
	for(var i = 1; i<appData.B.length; i++){
		if(!appData.roadMap[appData.B[i][3]]) appData.roadMap[appData.B[i][3]] = {epics:{}, itr:{}};
		
		if(!appData.roadMap[appData.B[i][3]].epics[appData.B[i][16]]){
			appData.roadMap[appData.B[i][3]].epics[appData.B[i][16]] = {name:"N/A", features:{}, f_count:0, s_count:0, s_points:0};
		}
		if(!appData.roadMap[appData.B[i][3]].epics[appData.B[i][16]].features[appData.B[i][5]]){
			appData.roadMap[appData.B[i][3]].epics[appData.B[i][16]].f_count++;
			appData.roadMap[appData.B[i][3]].epics[appData.B[i][16]].features[appData.B[i][5]] = {name:appData.B[i][6], dependency:"N/A", team:"N/A", status:"N/A", itr:{}, s_count:0, s_points:0};
		}
		if(!appData.roadMap[appData.B[i][3]].epics[appData.B[i][16]].features[appData.B[i][5]].itr[appData.B[i][2]]){
			appData.roadMap[appData.B[i][3]].itr[appData.B[i][2]] = {};
			appData.roadMap[appData.B[i][3]].epics[appData.B[i][16]].features[appData.B[i][5]].itr[appData.B[i][2]] = {stories:{},s_count:0, s_points:0};
		}
		if(!appData.roadMap[appData.B[i][3]].epics[appData.B[i][16]].features[appData.B[i][5]].itr[appData.B[i][2]].stories[appData.B[i][0]]){
			var p = isNaN(parseInt(appData.B[i][14])) ? 0 : parseInt(appData.B[i][14]);
			appData.roadMap[appData.B[i][3]].epics[appData.B[i][16]].s_count++;
			appData.roadMap[appData.B[i][3]].epics[appData.B[i][16]].s_points = appData.roadMap[appData.B[i][3]].epics[appData.B[i][16]].s_points + p;
			appData.roadMap[appData.B[i][3]].epics[appData.B[i][16]].features[appData.B[i][5]].itr[appData.B[i][2]].s_count++;
			
			appData.roadMap[appData.B[i][3]].epics[appData.B[i][16]].features[appData.B[i][5]].s_count++;
			appData.roadMap[appData.B[i][3]].epics[appData.B[i][16]].features[appData.B[i][5]].s_points = appData.roadMap[appData.B[i][3]].epics[appData.B[i][16]].features[appData.B[i][5]].s_points + p;
			
			appData.roadMap[appData.B[i][3]].epics[appData.B[i][16]].features[appData.B[i][5]].itr[appData.B[i][2]].s_points = appData.roadMap[appData.B[i][3]].epics[appData.B[i][16]].features[appData.B[i][5]].itr[appData.B[i][2]].s_points + p;
			appData.roadMap[appData.B[i][3]].epics[appData.B[i][16]].features[appData.B[i][5]].itr[appData.B[i][2]].stories[appData.B[i][0]] = {title:appData.B[i][1], priority:appData.B[i][10], owner:appData.B[i][11], status:appData.B[i][12], estimate:p};
		}
	}
	genRoadmapTables();
}

function genRoadmapTables(){
	$("#Roadmap").html("");
	var html = "";
	$.each(appData.roadMap, function(key1, v1){
		html+= "<table class='w3-medium w3-center w3-padding'>";
		html+= "<tr><td colspan='10' style='background-color:orange;'>"+key1+"</td></tr>";
		$.each(v1.epics, function(key2, v2){
			html+= "<tr><td colspan='10' style='background-color:MediumSpringGreen;'>"+key2+" : "+v2.name+" | Features :"+v2.f_count+" | Stories :"+v2.s_count+" | Story Points :"+v2.s_points+"</td></tr>";
			$.each(v2.features, function(key4, v4){
				html+= "<tr style='background-color:LightYellow;'><td>Epic ID</td><td>Feature ID</td><td>Feature Name</td><td>Dependency</td><td>Team</td><td>Status</td>";
				$.each(v4.itr, function(key3, v3){
					html+="<td>"+key3+"<br>Stories :"+v4.itr[key3].s_count+" | Points :"+v4.itr[key3].s_points+"</td>";
				});
				html+="</tr>";
				html+= "<tr><td>"+key2+"</td><td>"+key4+"</td><td>"+v4.name+"<br> Stories :"+v4.s_count+" | Points :"+v4.s_points+"</td><td>"+v4.dependency+"</td><td>"+v4.team+"</td><td>"+v4.status+"</td>";
				$.each(v4.itr, function(key5, v5){
					html+= "<td><table class='w3-padding'>";
					$.each(v5.stories, function(key6, v6){
						html+= "<tr><td>"+key6+" :"+v6.title+"<br>"+v6.status+"<br>"+v6.owner+"<br>"+v6.estimate+" Points</td></tr>";
					});
					html+= "</table></td>";
				});	
				html+= "</tr>";
			});		
		});
		html+= "</table><hr>";
	});
	$("#Roadmap").html(html);
}


function initFeaturesTable(){
	appData.files.F = {};
	for(var i=0; i<appData['F'][0].length; i++){
		appData.files.F[appData['F'][0][i]] = {'i':i, 'v':true}
	}
}

function initBacklogsTable(){
	appData.files.B = {};
	for(var i=0; i<appData['B'][0].length; i++){
		appData.files.B[appData['B'][0][i]] = {'i':i, 'v':true}
	}
	var b_fid_col = appData.files.B['Portfolio Item ID'].i;
	var f_fid_col = appData.files.F['Feature ID'].i;
	var f_eid_col = appData.files.F['Epic ID'].i
	for(var i=0; i<appData['B'].length; i++){
		if(i==0) {
			appData['B'][i].push('Epic ID');
			appData.files.B['Epic ID'] = {'i':appData['B'][i].length - 1, 'v':true}
		}
		else {
			for(var j=1; j<appData['F'].length; j++){
				appData['B'][i][appData['B'][0].length-1] = "N/A";
				if(appData['B'][i][b_fid_col] == appData['F'][j][f_fid_col]){
					appData['B'][i][appData['B'][0].length-1] = appData['F'][j][f_eid_col];
					break;
				}	
			}
		}
	}
}

function genPDFReport(){
	toggleOverlay('block');	
	pdfMake.fonts = {
	   Ubuntu: {
		   normal: 'https://cdn.jsdelivr.net/fontsource/fonts/ubuntu-condensed@latest/latin-400-normal.woff',
		   bold: 'https://cdn.jsdelivr.net/fontsource/fonts/ubuntu-condensed@latest/latin-400-normal.woff'
	   },
	};	
	var content = [{text:'\n\n\n\n\nPI Roadmap Report\n\n\nGenerated On'+new Date().toLocaleString(), alignment:'center', style:'cover_page'}];
	content.push({toc: {title:{id:'repTOC', text: 'Table of Contents', style:'page_header', alignment:'center'}}, pageBreak: 'before'});	
	var rows = [];	
	$.each(appData.roadMap, function(key1, v1){
		content.push({text: 'Roadmap for '+key1, pageBreak: 'before', alignment:'center', style: 'page_header', tocItem:true, decoration: 'underline', linkToDestination:'repTOC'})
		rows = [];
		$.each(v1.epics, function(key2, v2){
			rows.push([{text: key2+" : "+v2.name+" | Features :"+v2.f_count+" | Stories :"+v2.s_count+" | Story Points :"+v2.s_points, fillColor: 'mediumspringgreen', colSpan: 5, alignment: 'center', tocItem:true, linkToDestination:'repTOC'}, {}, {}, {}, {}]);
			$.each(v2.features, function(key4, v4){
				rows.push([{text:'Feature Details: '+key4, fillColor: 'lightyellow', alignment: 'center', linkToDestination:'repTOC', decoration: 'underline'}]);
				$.each(v4.itr, function(key3, v3){
					rows[rows.length - 1].push({text:key3+"\nStories :"+v4.itr[key3].s_count+" | Points :"+v4.itr[key3].s_points, fillColor: 'lightyellow', alignment: 'center'});
				});
				rows.push([{text:[{text:'Epic: '+key2+'\n\n', alignment: 'left'}, {text:'Feature: '+key4+' | '+v4.name, alignment: 'left', tocItem:true, linkToDestination:'repTOC'}, {text:'\n\nDependency: '+v4.dependency+'\n\nTeam: '+v4.team+'\n\nStatus: '+v4.status+'\n\nStories: '+v4.s_count+' | Points: '+v4.s_points, alignment: 'left'}]}]);
				$.each(v4.itr, function(key5, v5){
					var bl = [];
					$.each(v5.stories, function(key6, v6){
						bl.push([key6+" :"+v6.title+"\n"+v6.status+"\n"+v6.owner+"\n"+v6.estimate+" Points"]);
					});					
					rows[rows.length - 1].push({table:{widths:['*'], body:bl}});
				});	
			});			
		});
		content.push({table:{widths:['*', '*', '*', '*', '*'], body:rows}});
	});

	var dd = {pageSize:'A4',pageOrientation:'landscape',watermark:{text:"Trisha's Toolbox",color:'#AF7AC5',opacity: 0.1,bold:true,italics:false},content:content,defaultStyle: {font:'Ubuntu',fontSize:10},styles:{cover_page:{fontSize:30,color:'indigo',bold:true}, page_header:{fontSize:20,color:'firebrick',bold:true}},footer: function(currentPage, pageCount){return {text: 'Page ' + currentPage.toString() + ' of ' + pageCount, alignment: 'center', fontSize: 10};}};
    pdfMake.createPdf(dd).download('PI_Report.pdf', function(){alert('Report successfully exported!'); toggleOverlay('none');});	
}