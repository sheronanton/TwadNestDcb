		var AREA='';
		var GRADE='';
		var ULBtype='';
		
		function init() {

    if (PRVLB == 'L') {
        // Default: ULB
        CONS = 'ULB';
        document.getElementById('box2').style.display = 'inline';
        document.getElementById('box1').style.display = 'none';
        document.getElementById('box3').style.display = 'none';

        // Corporation
        if (CTYPE == 1) {
            GRADE = 1;
            ULBtype = 2;
            callServer('ULBgrade');
        }
        // III grade Municipality
        else if (CTYPE == 2) {
            ULBtype = 3;
            GRADE = 3;
            callServer('ULBgrade');
        }
        // Municipality
        else if (CTYPE == 3) {
            ULBtype = 3;
            callServer('ULBgrade');
        }
        // Urban Town Panchayat
        else if (CTYPE == 4) {
            ULBtype = 1;
            AREA = '2';
            callServer('ULBgrade');
        }
        // Rural Town Panchayat
        else if (CTYPE == 5) {
            ULBtype = 1;
            AREA = '1';
            callServer('ULBgrade');
        }
        // Panchayat
        else if (CTYPE == 6) {
            CONS = 'Panch';
            document.getElementById('box1').style.display = 'inline';
            document.getElementById('box2').style.display = 'none';
        }
        else {
            callServer('ULBgrade');
        }

        if ((CTYPE == 4) || (CTYPE == 5)) {
            //callServer('areaType');
        }

        // ---------- Build header for PRVLB = 'L' ----------
        var tdNoData = document.getElementById('tdNoData');
        var tblHead  = document.getElementById('tblHead');
        var colHead  = document.getElementById('colHead');

        // Clear any existing headers
        while (colHead.firstChild) {
            colHead.removeChild(colHead.firstChild);
        }

        // For CTYPE 6 we have: Select, District, Block, Village Panchayat, LGD Code => 5 columns
        if (CTYPE == 6) {
            tdNoData.colspan = "5";
            tblHead.colspan  = "5";
        } else {
            // Existing layout: Select, District, Block?, CTYPEDESC
            tdNoData.colspan = "3";
            tblHead.colspan  = "3";
        }

        var th = document.createElement('th');
        th.innerHTML = 'Select';
        colHead.appendChild(th);

        th = document.createElement('th');
        th.innerHTML = 'District';
        colHead.appendChild(th);

        if (parseInt(CTYPE) <= 6 && parseInt(CTYPE) >= 6) {
            th = document.createElement('th');
            th.innerHTML = 'Block';
            colHead.appendChild(th);
        }

        th = document.createElement('th');
        th.innerHTML = CTYPEDESC;   // For CTYPE 6: 'Village Panchayat'
        colHead.appendChild(th);

        // Extra LGD header only for Panch
        if (CTYPE == 6) {
            th = document.createElement('th');
            th.innerHTML = 'LGD Code';
            colHead.appendChild(th);
        }
    }
    else if (PRVLB == 'P') {
        // ---------- Private ----------
        CONS = 'Private';

        document.getElementById('box1').style.display = 'none';
        document.getElementById('box2').style.display = 'none';
        document.getElementById('box3').style.display = 'inline';

        document.getElementById('tdNoData').colspan = "4";
        document.getElementById('tblHead').colspan  = "4";

        var colHead = document.getElementById('colHead');
        while (colHead.firstChild) {
            colHead.removeChild(colHead.firstChild);
        }

        var th = document.createElement('th');
        th.innerHTML = 'Select';
        colHead.appendChild(th);

        th = document.createElement('th');
        th.innerHTML = CTYPEDESC;
        colHead.appendChild(th);

        th = document.createElement('th');
        th.innerHTML = 'District';
        colHead.appendChild(th);
    }
    else {
        CONS = 'NOT SELECTED';
        alert("Please Select a Consumer Type");
        Exit();
        return;
    }

    search();
}
		

		function getTransport()
		{
		 var req = false;
		 try 
		 {
		       req= new ActiveXObject("Msxml2.XMLHTTP");
		 }
		 catch (e) 
		 {
		       try 
		       {
		            req = new ActiveXObject("Microsoft.XMLHTTP");
		       }
		       catch (e2) 
		       {
		            req = false;
		       }
		 }
		 if (!req && typeof XMLHttpRequest != 'undefined') 
		 {
		       req = new XMLHttpRequest();
		 }   
		 return req;
		} 
		
		
		function Exit()
		{
		   self.close();
		}
		
		
		function nullCheck(idx)
		{
			var ele = document.getElementById(idx);
			if(ele.value=="")
			{ 
			     return false;
			}
			return true;
		}

		

		function unloadCombo(eleName)
		{
			var cmb = document.getElementById(eleName);
			var len = cmb.length; 
			for(var i=len; i>1; i--)
			{
				cmb.removeChild(cmb.lastChild);
			}
		}

		
		
		function search()
		{
			//if(  (nullCheck('dis')||nullCheck('blk')||nullCheck('pan')&&(CONS=='Panch'))   ||   (nullCheck('ulb')&&(CONS=='ULB'))   ||   (nullCheck('priv')&&(CONS=='Private'))  )  //nullCheck('area')||nullCheck('ULBgrade')||
			{
				document.getElementById('cmbpage').value=1;
				callServer(CONS);
			}
/*			else
			{
				clearGrid();
			}
*/			

		}
		

		function callServer(param)
		{
			 if(param=='District')
             {
				 	url="../../../../../Consumer_Search?command=District";
                    var req=getTransport();
                    req.open("GET",url,true);        
                    req.onreadystatechange=function()
                    {
                       processResponse(req);
                    }   
                    req.send(null);
             }
			 
			 if(param=='Block')
             {
				 	var dis = document.getElementById('dis').value;
             		
				 	url="../../../../../Consumer_Search?command=Block&dis="+dis;
                    var req=getTransport();
                    req.open("GET",url,true);        
                    req.onreadystatechange=function()
                    {
                       processResponse(req);

                    }   
                    req.send(null);
             }
			 
			 else if(param == 'Panch') {
        var dis = document.getElementById('dis').value;
        var blk = document.getElementById('blk').value;
        var pan = (document.getElementById('pan').value).toLowerCase();
        document.getElementById('divcmbpage').style.display = "inline";
        document.getElementById('divpage').style.display = "inline";
        var page = document.getElementById('cmbpage').value;
        var limit = document.getElementById('cmbpagination').value;
        url = "../../../../../Consumer_Search?command=Panch&dis=" + dis + "&blk=" + blk + "&pan=" + pan + "&page=" + page + "&limit=" + limit;
        req = getTransport();
        req.open("GET", url, true);
        req.onreadystatechange = function() {
            processResponse(req);
            // Refetch habitations after Panchayat listing
            //callServer('RdHabitation');
        };
        req.send(null);
    }
    else if(param == 'RdHabitation') {
        var dis = document.getElementById('dis').value;
        var blk = document.getElementById('blk').value;
        url = "../../../../../Consumer_Search?command=RdHabitation&dis=" + dis + "&blk=" + blk;
        req = getTransport();
        req.open("GET", url, true);
        req.onreadystatechange = function() { processResponse(req); };
        req.send(null);
    }
             
             
             
			 
/*	 		 if(param=='areaType')
             {
				 	url="../../../../../Consumer_Search?command=areaType";
                    var req=getTransport();
                    req.open("GET",url,true);        
                    req.onreadystatechange=function()
                    {
                       processResponse(req);
                    }   
                    req.send(null);
             }
*/			 
			 if(param=='ULBgrade')
             {
				 	url="../../../../../Consumer_Search?command=ULBgrade";
                    var req=getTransport();
                    req.open("GET",url,true);        
                    req.onreadystatechange=function()
                    {
                       processResponse(req);
                    }   
                    req.send(null);
             }
			 
			 else if(param=='ULB')
             {
				 	var area = AREA;//document.getElementById('area').value;
				 	var ULBgrade = document.getElementById('ULBgrade').value; // =GRADE;
				 	var ulb = (document.getElementById('ulb').value).toLowerCase();
				 	
				 	document.getElementById('divcmbpage').style.display = "inline";
				 	document.getElementById('divpage').style.display = "inline";

				 	var page = document.getElementById('cmbpage').value;
				 	var limit = document.getElementById('cmbpagination').value;
             		
				 	url="../../../../../Consumer_Search?command=ULB&ULBtype="+ULBtype+"&area="+area+"&ULBgrade="+ULBgrade+"&ulb="+ulb+"&page="+page+"&limit="+limit;
				 	var req=getTransport();
                    req.open("GET",url,true);        
                    req.onreadystatechange=function()
                    {
                       processResponse(req);
                    }   
                    req.send(null);
             }
			 
			 
			 else if(param=='Private')
             {
				 	var priv = (document.getElementById('priv').value).toLowerCase();
				 	
				 	document.getElementById('divcmbpage').style.display = "inline";
				 	document.getElementById('divpage').style.display = "inline";

				 	var page = document.getElementById('cmbpage').value;
				 	var limit = document.getElementById('cmbpagination').value;
             		
				 	url="../../../../../Consumer_Search?command=Private&priv="+priv+"&ctype="+CTYPE+"&page="+page+"&limit="+limit;
				 	
                 	var req=getTransport();
                    req.open("GET",url,true);        
                    req.onreadystatechange=function()
                    {
                       processResponse(req);
                    }   
                    req.send(null);
             }
			 

		}
		

		
		
		
		
		
   		function processResponse(req)
   		{
            if(req.readyState==4)
            {
                if(req.status==200)
                {
                    var baseResponse = req.responseXML.getElementsByTagName('response')[0];
                    var cmd = baseResponse.getElementsByTagName('command')[0].firstChild.nodeValue;
                    var flag = baseResponse.getElementsByTagName('flag')[0].firstChild.nodeValue;
                    
                    if(cmd == 'District')
                    {
                    		document.getElementById('cmbpage').value=1;
                    		
                    		var cmb = document.getElementById('dis');
                            
                            var row = baseResponse.getElementsByTagName('row');
                            var len = row.length;
                            unloadCombo('dis');

                            for(var i=0; i<len; i++)
                            {
                                    var dis = row[i].getElementsByTagName('dis')[0].firstChild.nodeValue;
                                    var district = row[i].getElementsByTagName('district')[0].firstChild.nodeValue;

                                    var opt = document.createElement('option');
                                    opt.value = dis;
                                    opt.innerHTML = district;
                                    
                                    cmb.appendChild(opt);
                            }
                    }                    
                    
                    else if(cmd == 'Block')
                    {
                    		document.getElementById('cmbpage').value=1;
                    		
                    		var cmb = document.getElementById('blk');
                            
                            var row = baseResponse.getElementsByTagName('row');
                            var len = row.length;
                            unloadCombo('blk');

                            for(var i=0; i<len; i++)
                            {
                                    var pan = row[i].getElementsByTagName('blk')[0].firstChild.nodeValue;
                                    var panch = row[i].getElementsByTagName('block')[0].firstChild.nodeValue;

                                    var opt = document.createElement('option');
                                    opt.value = pan;
                                    opt.innerHTML = panch;
                                    
                                    cmb.appendChild(opt);
                            }
                            
                            search();
                    }                    

                    else if(cmd == 'Panch')
                    {
	            		
                    	    var tbody = document.getElementById('tbody');
                            
                            var row = baseResponse.getElementsByTagName('row');
                            var len = row.length;

                            unloadChildren('tbody');
                            unloadCombo('cmbpage');
                            
                            var page = new Number(baseResponse.getElementsByTagName('page')[0].firstChild.nodeValue);
                            var totpg = new Number(baseResponse.getElementsByTagName('total')[0].firstChild.nodeValue);
                            

                            
                            
                            /****** Load 'Page No.' Combo & 'Total Pages' **********/ 
                            
                            document.getElementById('divpage').innerHTML = totpg;
                            var cmbpage = document.getElementById('cmbpage');
                            
                            for(var i=2; i<=totpg; i++)
                            {
                            	var opt = document.createElement('option');
                            	opt.value = i;
                            	opt.innerHTML = i;
                            	cmbpage.appendChild(opt);
                            }
                            cmbpage.value = page; 
                            
                            /*******************************************************/
                            
                            
	          
                            
                            
                            /************* 'Next' & 'Previous' links **************/
                            
                            if(page<totpg)
                            {
                            	document.getElementById('divnext').style.display = 'inline';
                            }
                            else
                            {
                            	document.getElementById('divnext').style.display = 'none';
                            }
                            
                            
                            if(page>1)
                            {
                            	document.getElementById('divpre').style.display = 'inline';
                            }
                            else
                            {
                            	document.getElementById('divpre').style.display = 'none';
                            }
                            /*******************************************************/
                            
                            
                            
                            document.getElementById('nodata').style.display="none";     // Hide 'No Data Found' msg
                            
                            if(totpg==0)
                            {
                            	hidePag();
                            }
                            
                            
                            for(var i=0; i<len; i++)
                            {
	                            	var dis = row[i].getElementsByTagName('dis')[0].firstChild.nodeValue;
	                            	var blk = row[i].getElementsByTagName('blk')[0].firstChild.nodeValue;
                            		var pan = row[i].getElementsByTagName('pan')[0].firstChild.nodeValue;
                            		var district = row[i].getElementsByTagName('district')[0].firstChild.nodeValue;
                            		var block = row[i].getElementsByTagName('block')[0].firstChild.nodeValue;
                            		var panch = row[i].getElementsByTagName('panch')[0].firstChild.nodeValue;
                            		var lgd      = row[i].getElementsByTagName('lgd')[0].firstChild.nodeValue;
                                    var tr = document.createElement('tr');
                                    tr.id = pan;
                                    
                                    var tdRad = document.createElement('td');
                                    var rad;
                                    if (window.navigator.appName.toLowerCase().indexOf("netscape") == -1)
                                    {

                                    	rad = document.createElement("<input type='radio' name='rad' id='rad' value="+pan+" onclick='dupCheck(this.value);'>");
                                    	
                                    	
                                    }else
                                    {
                                    	
                                    	rad = document.createElement('input');
                                        rad.name = 'rad';
                                        rad.type = 'radio';
                                        rad.value = pan;
                                    	//rad.onclick='dupCheck(this.value)';
                                        rad.setAttribute('onclick',"dupCheck(" + pan + ")");
                                        
                                    }

                                    tdRad.appendChild(rad);
                                    tr.appendChild(tdRad);

                                    var tdDis = document.createElement('td');
                                    tdDis.innerHTML = dis;
                                    tdDis.style.display='none';
                                    tr.appendChild(tdDis);

                                    var tdBlk = document.createElement('td');
                                    tdBlk.innerHTML = blk;
                                    tdBlk.style.display='none';
                                    tr.appendChild(tdBlk);

                                    var tdPan = document.createElement('td');
                                    tdPan.innerHTML = pan;
                                    tdPan.style.display='none';
                                    tr.appendChild(tdPan);
                                    
                                    var tdDistrict = document.createElement('td');
                                    tdDistrict.innerHTML = district;
                                    tr.appendChild(tdDistrict);
                                    
                                    var tdBlock = document.createElement('td');
                                    tdBlock.innerHTML = block;
                                    tr.appendChild(tdBlock);

                                    var tdPanch = document.createElement('td');
                                    tdPanch.innerHTML = panch;
                                    tr.appendChild(tdPanch);
                                    
                                    
                                     // 5: LGD code  <<< NEW CELL
								      var tdLgd =  document.createElement('td');
								      tdLgd.innerHTML = lgd;
								      tr.appendChild(tdLgd);
                                    
                                    tbody.appendChild(tr);
                            }
                    }
                    
                                           if (cmd === "RdHabitation") {
            var rows = xml.getElementsByTagName("row");
            var rdDropdown = document.getElementById("rd_panchayat");
            rdDropdown.innerHTML = '<option value="">----Select Habitation----</option>';
            for (var i = 0; i < rows.length; i++) {
                var codeNode = rows[i].getElementsByTagName("rd_panchayat_code")[0];
                var nameNode = rows[i].getElementsByTagName("rd_panchayat_name")[0];
                var code = codeNode ? (codeNode.textContent || codeNode.firstChild.nodeValue) : "";
                var name = nameNode ? (nameNode.textContent || nameNode.firstChild.nodeValue) : "";
                var option = document.createElement("option");
                option.value = code;
                option.text = name;
                rdDropdown.appendChild(option);
            }
            return;
        }
                    
   /*                 
                    else if(cmd == 'areaType')
                    {
                    		document.getElementById('cmbpage').value=1;
                    		
                    		var cmb = document.getElementById('area');
                            
                            var row = baseResponse.getElementsByTagName('row');
                            var len = row.length;
                            unloadCombo('area');
                            for(var i=0; i<len; i++)
                            {
                                    var tid = row[i].getElementsByTagName('aid')[0].firstChild.nodeValue;
                                    var tdesc = row[i].getElementsByTagName('adesc')[0].firstChild.nodeValue;

                                    var opt = document.createElement('option');
                                    opt.value = tid;
                                    opt.innerHTML = tdesc;
                                    
                                    cmb.appendChild(opt);
                            }
                            if(CTYPE==4)
                            {
                            	cmb.value=1;
                            }
                            else if(CTYPE==5)
                            {
                            	cmb.value=2;
                            }
                            search();
                    }     
*/
                    
                    else if(cmd == 'ULBgrade')
                    {
                    		document.getElementById('cmbpage').value=1;
                    		
                    		var cmb = document.getElementById('ULBgrade');
                            
                            var row = baseResponse.getElementsByTagName('row');
                            var len = row.length;
                            unloadCombo('ULBgrade');

                            for(var i=0; i<len; i++)
                            {
                                    var gid = row[i].getElementsByTagName('gid')[0].firstChild.nodeValue;
                                    var gdesc = row[i].getElementsByTagName('gdesc')[0].firstChild.nodeValue;
                                    
                                    
                                    /******************************************************
                                     * 	If Municipality - grade III should not be displayed
                                     ******************************************************/
                                    if(CTYPE==3) // Municipality
                                    {
                                    	document.getElementById('ULBgrade').disabled = false;
                                    	if(gid==3)
                                    	{
                                    		continue; // Skip Grade-III
                                    	}
                                    }
                                    else if((CTYPE>6)||(CTYPE==4)||(CTYPE==5))
                                    {
                                    	document.getElementById('ULBgrade').disabled = false;
                                    }
                                    else
                                    {
                                    	document.getElementById('ULBgrade').disabled = true;
                                    }
                                    /******************************************************/
                                    
                                    
                                    var opt = document.createElement('option');
                                    opt.value = gid;
                                    opt.innerHTML = gdesc;
                                    
                                    cmb.appendChild(opt);
                            }
                            
                            
                            if(CTYPE==1)
                            {
                            	cmb.value=1;
                            }
                            else if(CTYPE==2)
                            {
                            	cmb.value=3;
                            }
                            else if(CTYPE==4) // UTP
                            {
                            	//cmb.value=5; // Spl grade  (UTP Spl)
                            }
                            else if(CTYPE==5) // RTP
                            {
                            	//cmb.value=5; // Spl grade  (RTP Spl)
                            }
                            
                            search();
                    }                 
    
                    else if(cmd == 'ULB')
                    {
               		
                    		
                    		
                            var tbody = document.getElementById('tbody');
                            
                            var row = baseResponse.getElementsByTagName('row');
                            var len = row.length;

                            unloadChildren('tbody');
                            unloadCombo('cmbpage');
                            
                            var page = new Number(baseResponse.getElementsByTagName('page')[0].firstChild.nodeValue);
                            var totpg = new Number(baseResponse.getElementsByTagName('total')[0].firstChild.nodeValue);
                            

                            
                            
                            /****** Load 'Page No.' Combo & 'Total Pages' **********/ 
                            
                            document.getElementById('divpage').innerHTML = totpg;
                            var cmbpage = document.getElementById('cmbpage');
                            
                            for(var i=2; i<=totpg; i++)
                            {
                            	var opt = document.createElement('option');
                            	opt.value = i;
                            	opt.innerHTML = i;
                            	cmbpage.appendChild(opt);
                            }
                            cmbpage.value = page; 
                            
                            /*******************************************************/
                            
                            
                            
                            
                            
                            /************* 'Next' & 'Previous' links **************/
                            
                            if(page<totpg)
                            {
                            	document.getElementById('divnext').style.display = 'inline';
                            }
                            else
                            {
                            	document.getElementById('divnext').style.display = 'none';
                            }
                            
                            
                            if(page>1)
                            {
                            	document.getElementById('divpre').style.display = 'inline';
                            }
                            else
                            {
                            	document.getElementById('divpre').style.display = 'none';
                            }
                            /*******************************************************/
                            
                            
                            
                            document.getElementById('nodata').style.display="none";     // Hide 'No Data Found' msg
                            
                            if(totpg==0)
                            {
                            	hidePag();
                            }
                            
                            
                            
                            for(var i=0; i<len; i++)
                            {
                            		var dis = baseResponse.getElementsByTagName('dis')[i].firstChild.nodeValue;
                            		var district = baseResponse.getElementsByTagName('district')[i].firstChild.nodeValue;
                                    var ulb = row[i].getElementsByTagName('ulb')[0].firstChild.nodeValue;
                                    var urbanlb = row[i].getElementsByTagName('urbanlb')[0].firstChild.nodeValue;

                                    var tr = document.createElement('tr');
                                    tr.id = ulb;

                                    
                                    var tdRad = document.createElement('td');
                                    var rad;
                                    if (window.navigator.appName.toLowerCase().indexOf("netscape") == -1)
                                    {

                                    	rad = document.createElement("<input type='radio' name='rad' id='rad' value="+ulb+" onclick='dupCheck(this.value);'>");
                                    	
                                    }
                                    else
                                    {
                                    	
                                    	rad = document.createElement('input');
                                        rad.name = 'rad';
                                        rad.type = 'radio';
                                        rad.value = ulb;
                                    	//rad.onclick='dupCheck(this.value)';
                                        rad.setAttribute('onclick',"dupCheck(this.value)");
                                    }

                                    

                                    
                                    tdRad.appendChild(rad);
                                    tr.appendChild(tdRad);

                                    var tdulb = document.createElement('td');
                                    tdulb.innerHTML = ulb;
                                    tdulb.style.display="none";
                                    tr.appendChild(tdulb);

                                    var tddis = document.createElement('td');
                                    tddis.innerHTML = dis;
                                    tddis.style.display="none";
                                    tr.appendChild(tddis);

                                    var tddistrict = document.createElement('td');
                                    tddistrict.innerHTML = district;
                                    tr.appendChild(tddistrict);

                                    var tdurbanlb = document.createElement('td');
                                    tdurbanlb.innerHTML = urbanlb;
                                    tr.appendChild(tdurbanlb);
                                    
                                    tbody.appendChild(tr);
                            }
                    }
                 
                    
                    else if(cmd == 'Private')
                    {
               		
                    	
                            var tbody = document.getElementById('tbody');
                            
                            var row = baseResponse.getElementsByTagName('row');
                            var len = row.length;
                            
                            unloadChildren('tbody');
                            unloadCombo('cmbpage');

                            var page = new Number(baseResponse.getElementsByTagName('page')[0].firstChild.nodeValue);
                            var totpg = new Number(baseResponse.getElementsByTagName('total')[0].firstChild.nodeValue);
                            
                            
                            
                            
                            /****** Load 'Page No.' Combo & 'Total Pages' **********/ 
                            
                            document.getElementById('divpage').innerHTML = totpg;
                            var cmbpage = document.getElementById('cmbpage');
                                                    
                            for(var i=2; i<=totpg; i++)
                            {
                            	var opt = document.createElement('option');
                            	opt.value = i;
                            	opt.innerHTML = i;
                            	cmbpage.appendChild(opt);
                            }
                            cmbpage.value = page; 
                            
                            /*******************************************************/
                            
                            
                            
                            
                            /************* 'Next' & 'Previous' links **************/
                            
                            if(page<totpg)
                            {
                            	document.getElementById('divnext').style.display = 'inline';
                            }
                            else
                            {
                            	document.getElementById('divnext').style.display = 'none';
                            }
                            
                            
                            if(page>1)
                            {
                            	document.getElementById('divpre').style.display = 'inline';
                            }
                            else
                            {
                            	document.getElementById('divpre').style.display = 'none';
                            }
                            /*******************************************************/
                            
                            
                            
                            document.getElementById('nodata').style.display="none";     // Hide 'No Data Found' msg
                            if(totpg==0)
                            {
                            	hidePag();
                            }
                            
                            
                            for(var i=0; i<len; i++)
                            {
                                    var pid = row[i].getElementsByTagName('pid')[0].firstChild.nodeValue;
                                    var pdesc = row[i].getElementsByTagName('pdesc')[0].firstChild.nodeValue;
                                    //var grp = row[i].getElementsByTagName('grp')[0].firstChild.nodeValue;
                                    var dis = row[i].getElementsByTagName('dis')[0].firstChild.nodeValue;
                                    var district = row[i].getElementsByTagName('district')[0].firstChild.nodeValue;

                                    var tr = document.createElement('tr');
                                    tr.id = pid;  //ulb
                                    
                                    var tdRad = document.createElement('td');
                                    var rad;
                                    if (window.navigator.appName.toLowerCase().indexOf("netscape") == -1)
                                    {

                                    	rad = document.createElement("<input type='radio' name='rad' id='rad' value="+pid+" onclick='dupCheck(this.value);'>");
                                    	
                                    	
                                    }else
                                    {
                                    	
                                    	rad = document.createElement('input');
                                        rad.name = 'rad';
                                        rad.type = 'radio';
                                        rad.value = pid;
                                    	//rad.onclick=dupCheck(this.value);
                                        rad.setAttribute('onclick',"dupCheck(this.value)");
                                    }
                                 
                                    tdRad.appendChild(rad);
                                    tr.appendChild(tdRad);

                                    var td;
                                    
                                    td = document.createElement('td');
                                    td.innerHTML = pid;
                                    td.style.display="none";
                                    tr.appendChild(td);

/*
                                    var tdgrp = document.createElement('td');
                                    tdgrp.innerHTML = grp;
                                    tdgrp.style.display="none";
                                    tr.appendChild(tdgrp);
*/
                                    var td = document.createElement('td');
                                    td.innerHTML = pdesc;
                                    tr.appendChild(td);

                                    var td = document.createElement('td');
                                    td.innerHTML = dis;
                                    td.style.display="none";
                                    tr.appendChild(td);

                                    td = document.createElement('td');
                                    td.innerHTML = district;
                                    tr.appendChild(td);
                                    
                                    
                                    tbody.appendChild(tr);
                            }
                    }
                    
                    else if(cmd == 'Duplicate')
                    {
                    	if(flag=='duplicate')
                    	{
                    		alert('Beneficiary already Exists. Please select another Beneficiary.');
                    	}
                    }                    
                    

                    
                    
                    
                    
                }
            }
   		}
   		

		
		
		  

		function clearGrid()
		{
			unloadChildren('tbody');
			hidePag();
		}

		
		
		function unloadChildren(eleName)
		{
			var prnt = document.getElementById(eleName);
			var len = prnt.childNodes.length;
			for(var i=len; i>0; i--)
			{
				prnt.removeChild(prnt.lastChild);
			}
		}
	
	
	   
	   function prev()
	   {
		   var page = document.getElementById('cmbpage');
		   
		   if(new Number(page.value>1))
		   {
			   (page.value)--;
		   }
		   callServer(CONS);
	   }
	

	   
	   function next()
	   {
		   var page = document.getElementById('cmbpage');
		   var totpg = new Number(document.getElementById('divpage').firstChild.nodeValue);
		   if(new Number(page.value<totpg))
		   {
			   (page.value)++;
		   }
		   callServer(CONS);
	   }
	   
	   
	   
	   function hidePag()
	   {
		   document.getElementById('divnext').style.display="none";
		   document.getElementById('divpre').style.display="none";
		   document.getElementById('divpage').style.display="none";
		   document.getElementById('divcmbpage').style.display="none";
		   
		   document.getElementById('nodata').style.display="inline";     // Display 'No Data Found' msg
	   }
	   
	   
	   
		function done_new()
{
    var rad = document.getElementsByName('rad');
    var len = rad.length;
    for (var i = 0; i < len; i++)
    {
        if (rad[i].checked == true)
        {
            var selRow = document.getElementById(rad[i].value);
            
            var cname   = selRow.childNodes[6].lastChild.nodeValue;
// LGD Code
var lgdCode = selRow.childNodes[7].lastChild.nodeValue;

            var dis  = 0;
            var blk  = 0;
            var pan  = 0;
            var priv = 0;
            var ulb  = 0;

            if (CONS == 'Panch')
            {
                dis = selRow.childNodes[1].lastChild.nodeValue;
                blk = selRow.childNodes[2].lastChild.nodeValue;
                pan = selRow.childNodes[3].lastChild.nodeValue;
                // e.g. if LGD is at index 4 for Panch:
                                //lgdCode = selRow.childNodes[7].lastChild.nodeValue; // LGD code (234263)

			         }
            else if (CONS == 'ULB')
            {
                ulb = selRow.childNodes[1].lastChild.nodeValue;
                dis = selRow.childNodes[2].lastChild.nodeValue;
                                cname = selRow.childNodes[3].lastChild.nodeValue; // ULB / Town Panchayat name

                // e.g. if LGD is at index 3 for ULB:
                // lgdCode = selRow.childNodes[3].lastChild.nodeValue;
            }
            else if (CONS == 'Private')
            {
                priv  = selRow.childNodes[1].lastChild.nodeValue;
                dis   = selRow.childNodes[3].lastChild.nodeValue;
                cname = selRow.childNodes[2].lastChild.nodeValue;
            }

            opener.doParent(cname, dis, blk, pan, priv, ulb, lgdCode);  // pass LGD
            Exit();
        }
    }
}
		function done()
{
    var rad = document.getElementsByName('rad');
    var len = rad.length;
    for (var i = 0; i < len; i++)
    {
        if (rad[i].checked == true)
        {
            var selRow = document.getElementById(rad[i].value);

            var dis  = 0;
            var blk  = 0;
            var pan  = 0;
            var priv = 0;
            var ulb  = 0;
            var lgdCode = "";
            var cname = "";

            if (CONS == 'Panch')
            {
                // Panch layout:
                // 0 radio, 1 dis, 2 blk, 3 pan, 4 district, 5 block, 6 panch, 7 lgd
                dis     = selRow.childNodes[1].lastChild.nodeValue;
                blk     = selRow.childNodes[2].lastChild.nodeValue;
                pan     = selRow.childNodes[3].lastChild.nodeValue;
                cname   = selRow.childNodes[6].lastChild.nodeValue; // Village Panchayat
                lgdCode = selRow.childNodes[7].lastChild.nodeValue; // LGD code
            }
            else if (CONS == 'ULB')
            {
                // ULB (Town Panchayat) layout:
                // 0 radio, 1 ulb id (hidden), 2 dis (hidden), 3 ulb name, 4 district name
                ulb   = selRow.childNodes[1].lastChild.nodeValue; // ulb id
                dis   = selRow.childNodes[2].lastChild.nodeValue; // district code
                cname = selRow.childNodes[4].lastChild.nodeValue; // town panchayat name
                // no LGD for ULB
            }
            else if (CONS == 'Private')
            {
                // Private layout (original):
                // 0 radio, 1 pid (hidden), 2 name, 3 dis (hidden), 4 district name
                priv  = selRow.childNodes[1].lastChild.nodeValue;
                dis   = selRow.childNodes[3].lastChild.nodeValue;
                cname = selRow.childNodes[2].lastChild.nodeValue;
            }

            opener.doParent(cname, dis, blk, pan, priv, ulb, lgdCode);
            Exit();
        }
    }
}
		
		
		function dupCheck(rid)
		{
			var selRow = document.getElementById(rid);
			
			var dis = 0;
			var blk = 0;
			var pan = 0;
			
			var priv = 0;
			//var grp = 0;
			
			var ulb = 0;

			if(CONS=='Panch')
			{
				dis = selRow.childNodes[1].lastChild.nodeValue;
				blk = selRow.childNodes[2].lastChild.nodeValue;
				pan = selRow.childNodes[3].lastChild.nodeValue;
			}
			else if(CONS=='ULB')
			{
				ulb = selRow.childNodes[1].lastChild.nodeValue;
				dis = selRow.childNodes[2].lastChild.nodeValue;
			}
			else if(CONS=='Private')
			{
				priv = selRow.childNodes[1].lastChild.nodeValue;
				dis = selRow.childNodes[3].lastChild.nodeValue;
			}
			
		 	url="../../../../../Consumer_Search?command=Duplicate&pan="+pan+"&priv="+priv+"&ulb="+ulb;
            var req=getTransport();
            req.open("GET",url,true);        
            req.onreadystatechange=function()
            {
               processResponse(req);
            }
            req.send(null);


		}
 	   