<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
pageEncoding="ISO-8859-1"%>
<%@ page import="java.util.Calendar"%>
<%@ page import="java.sql.Connection"%>
<%@ page import="java.sql.PreparedStatement"%>
<%@ page import="java.sql.ResultSet"%>
<%@ page import="java.sql.ResultSetMetaData"%>
<%@ page import="Servlets.PMS.PMS1.DCB.servlets.Controller"%>

<%
request.setCharacterEncoding("ISO-8859-1");

String action = request.getParameter("action");
String ajaxDistrictId = request.getParameter("districtId");

// ===== selected office from dropdown =====
String selectedOfficeId = request.getParameter("officeId");
if (selectedOfficeId == null || selectedOfficeId.equals("")) {
    selectedOfficeId = "0";
}
// =========================================

if ("getBlocks".equals(action)) {
    Connection ajaxCon = null;
    PreparedStatement ajaxPs = null;
    ResultSet ajaxRs = null;

    try {
        out.print("<!-- districtId=" + ajaxDistrictId + " -->");

        Controller ajaxObj = new Controller();
        ajaxCon = ajaxObj.con();

        out.print("<option value='0'>Select Block</option>");

        String ajaxSql =
            "SELECT block_sno, block_name " +
            "FROM com_mst_blocks " +
            "WHERE district_code = ? " +
            "ORDER BY block_name";

        ajaxPs = ajaxCon.prepareStatement(ajaxSql);
        ajaxPs.setInt(1, Integer.parseInt(ajaxDistrictId));
        ajaxRs = ajaxPs.executeQuery();

        while (ajaxRs.next()) {
            out.print("<option value='" + ajaxRs.getString("block_sno") + "'>" +
                      ajaxRs.getString("block_name") + "</option>");
        }
    } catch (Exception e) {
        e.printStackTrace();
        out.print("<option value='0'>Error : " + e.getMessage() + "</option>");
    } finally {
        try { if (ajaxRs != null) ajaxRs.close(); } catch (Exception ex) {}
        try { if (ajaxPs != null) ajaxPs.close(); } catch (Exception ex) {}
        try { if (ajaxCon != null) ajaxCon.close(); } catch (Exception ex) {}
    }
    return;
}

String exportToExcel = request.getParameter("exportToExcel");
boolean isExcel = "YES".equalsIgnoreCase(exportToExcel);

if (isExcel) {
    response.reset();
    response.setContentType("application/vnd.ms-excel; charset=ISO-8859-1");
    response.setHeader("Content-Disposition", "attachment; filename=DCB_Beneficiary_Monthwise_Report.xls");
}
%>

<%!
public String formatColumnName(String colName) {
    if (colName == null || colName.trim().equals("")) {
        return "";
    }

    String[] parts = colName.toLowerCase().split("_");
    StringBuffer sb = new StringBuffer();

    for (int i = 0; i < parts.length; i++) {
        if (parts[i].length() > 0) {
            sb.append(parts[i].substring(0, 1).toUpperCase());
            if (parts[i].length() > 1) {
                sb.append(parts[i].substring(1));
            }
            if (i < parts.length - 1) {
                sb.append(" ");
            }
        }
    }
    return sb.toString();
}
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>DCB Beneficiary Monthwise Report</title>
<link href='../../../../../../css/AME.css' rel='stylesheet' media='screen' />
<script type="text/javascript" src="../scripts/collection_amount_new.js"></script>
<script type="text/javascript" src="../../scripts/cellcreate.js"></script>
<script type="text/javascript" src="../../scripts/dcbvalidation.js"></script>

<script type="text/javascript">

function exportExcel() {
    var fromYear = document.getElementById("fromYear").value;
    var fromMonth = document.getElementById("fromMonth").value;
    var toYear = document.getElementById("toYear").value;
    var toMonth = document.getElementById("toMonth").value;
    var beneficiaryType = document.getElementById("beneficiaryType").value;
    var beneficiaryId = document.getElementById("beneficiaryId").value;
    var districtId = document.getElementById("districtId") ? document.getElementById("districtId").value : "0";
    var blockId = document.getElementById("blockId") ? document.getElementById("blockId").value : "0";
    var officeId = document.getElementById("officeId") ? document.getElementById("officeId").value : "0";

    // optional: basic validation like funcmddata()
    if (fromYear == "0" || fromMonth == "0" || toYear == "0" || toMonth == "0") {
        alert("Please select all period fields");
        return false;
    }
    if (beneficiaryType == "0") {
        alert("Please select Beneficiary Type");
        return false;
    }
    if (beneficiaryType == "6" && districtId == "0") {
        alert("Please select District");
        return false;
    }
    if (beneficiaryType == "6" && blockId == "0") {
        alert("Please select Block");
        return false;
    }
    if (beneficiaryId == "0") {
        alert("Please select Beneficiary");
        return false;
    }

    var url =
        "BeneficiaryMonthwise.jsp?exportToExcel=YES" +
        "&mode=print" +
        "&fromYear=" + encodeURIComponent(fromYear) +
        "&fromMonth=" + encodeURIComponent(fromMonth) +
        "&toYear=" + encodeURIComponent(toYear) +
        "&toMonth=" + encodeURIComponent(toMonth) +
        "&beneficiaryType=" + encodeURIComponent(beneficiaryType) +
        "&beneficiaryId=" + encodeURIComponent(beneficiaryId) +
        "&districtId=" + encodeURIComponent(districtId) +
        "&blockId=" + encodeURIComponent(blockId) +
        "&officeId=" + encodeURIComponent(officeId);

    window.location.href = url;
}


function onOfficeChange() {
    // reset all filters when office changes
    document.getElementById("fromYear").value = "0";
    document.getElementById("fromMonth").value = "0";
    document.getElementById("toYear").value = "0";
    document.getElementById("toMonth").value = "0";

    document.getElementById("beneficiaryType").value = "0";

    if (document.getElementById("districtId")) {
        document.getElementById("districtId").value = "0";
    }
    if (document.getElementById("blockId")) {
        document.getElementById("blockId").innerHTML = "<option value='0'>Select Block</option>";
    }
    if (document.getElementById("beneficiaryId")) {
        document.getElementById("beneficiaryId").innerHTML = "<option value='0'>- - - Select - - -</option>";
    }

    // submit with only office and default filters
    document.pms_dcb_ledger_report.submit();
} 

function getBlocks(districtId) {
    var block = document.getElementById("blockId");

    if (districtId == null || districtId == "" || districtId == "0") {
        block.innerHTML = "<option value='0'>Select Block</option>";
        return;
    }

    block.innerHTML = "<option value='0'>Loading...</option>";

    var xhr = new XMLHttpRequest();
    xhr.open("GET", "BeneficiaryMonthwise.jsp?action=getBlocks&districtId=" + encodeURIComponent(districtId), true);

    xhr.onreadystatechange = function() {
        if (xhr.readyState == 4) {
            if (xhr.status == 200) {
                block.innerHTML = xhr.responseText;
            } else {
                block.innerHTML = "<option value='0'>Unable to load blocks</option>";
            }
        }
    };

    xhr.send();
}

function onBeneficiaryTypeChange() {
    var benType = document.getElementById("beneficiaryType").value;
    var districtRow = document.getElementById("districtRow");
    var blockRow = document.getElementById("blockRow");

    if (benType == "6") {
        districtRow.style.display = "";
        blockRow.style.display = "";
    } else {
        districtRow.style.display = "none";
        blockRow.style.display = "none";
        if (document.getElementById("districtId")) {
            document.getElementById("districtId").value = "0";
        }
        if (document.getElementById("blockId")) {
            document.getElementById("blockId").innerHTML = "<option value='0'>Select Block</option>";
        }
    }

    document.pms_dcb_ledger_report.submit();
}

function validatePeriod() {
    var fromYear = document.getElementById("fromYear").value;
    var fromMonth = document.getElementById("fromMonth").value;
    var toYear = document.getElementById("toYear").value;
    var toMonth = document.getElementById("toMonth").value;
    var beneficiaryType = document.getElementById("beneficiaryType").value;
    var beneficiaryId = document.getElementById("beneficiaryId").value;
    var districtId = document.getElementById("districtId") ? document.getElementById("districtId").value : "0";
    var blockId = document.getElementById("blockId") ? document.getElementById("blockId").value : "0";

    if (fromYear == "0") {
        alert("Please select From Year");
        document.getElementById("fromYear").focus();
        return false;
    }
    if (fromMonth == "0") {
        alert("Please select From Month");
        document.getElementById("fromMonth").focus();
        return false;
    }
    if (toYear == "0") {
        alert("Please select To Year");
        document.getElementById("toYear").focus();
        return false;
    }
    if (toMonth == "0") {
        alert("Please select To Month");
        document.getElementById("toMonth").focus();
        return false;
    }
    if (beneficiaryType == "0") {
        alert("Please select Beneficiary Type");
        document.getElementById("beneficiaryType").focus();
        return false;
    }
    if (beneficiaryType == "6" && districtId == "0") {
        alert("Please select District");
        document.getElementById("districtId").focus();
        return false;
    }
    if (beneficiaryType == "6" && blockId == "0") {
        alert("Please select Block");
        document.getElementById("blockId").focus();
        return false;
    }
    if (beneficiaryId == "0") {
        alert("Please select Beneficiary");
        document.getElementById("beneficiaryId").focus();
        return false;
    }

    var fromValue = parseInt(fromYear, 10) * 100 + parseInt(fromMonth, 10);
    var toValue = parseInt(toYear, 10) * 100 + parseInt(toMonth, 10);

    if (fromValue > toValue) {
        alert("From Period should not be greater than To Period");
        return false;
    }

    return true;
}

function funcmddata() {
    if (!validatePeriod()) {
        return false;
    }

    document.getElementById("mode").value = "print";
    document.pms_dcb_ledger_report.submit();
}

function refreshForm() {
    window.location.href = window.location.pathname;
}

function exitwindow() {
    window.close();
}
</script>

<style type="text/css">
.reportScroll {
    width: 95%;
    margin: 10px auto;
    max-height: 680px;
    overflow: auto;
    border: 1px solid #999;
}
.reportScroll table {
    width: 100%;
    border-collapse: collapse;
}
.reportScroll th {
    position: sticky;
    top: 0;
    background-color: #408099;
    color: white;
    z-index: 2;
}
</style>


</head>
<body>

<%
String userid = "0";
String Office_id = "";
String Office_Name = "";
Controller obj = null;
Connection con = null;

Calendar cal = Calendar.getInstance();
int month = cal.get(Calendar.MONTH) + 1;
int year = cal.get(Calendar.YEAR);

String selectedFromYear = request.getParameter("fromYear");
if (selectedFromYear == null || selectedFromYear.equals("")) selectedFromYear = "0";

String selectedFromMonth = request.getParameter("fromMonth");
if (selectedFromMonth == null || selectedFromMonth.equals("")) selectedFromMonth = "0";

String selectedToYear = request.getParameter("toYear");
if (selectedToYear == null || selectedToYear.equals("")) selectedToYear = "0";

String selectedToMonth = request.getParameter("toMonth");
if (selectedToMonth == null || selectedToMonth.equals("")) selectedToMonth = "0";

String selectedBeneficiaryType = request.getParameter("beneficiaryType");
if (selectedBeneficiaryType == null || selectedBeneficiaryType.equals("")) selectedBeneficiaryType = "0";

String selectedBeneficiaryId = request.getParameter("beneficiaryId");
if (selectedBeneficiaryId == null || selectedBeneficiaryId.equals("")) selectedBeneficiaryId = "0";

String selectedDistrictId = request.getParameter("districtId");
if (selectedDistrictId == null || selectedDistrictId.equals("")) selectedDistrictId = "0";

String selectedBlockId = request.getParameter("blockId");
if (selectedBlockId == null || selectedBlockId.equals("")) selectedBlockId = "0";

String mode = request.getParameter("mode");
if (mode == null) mode = "";

try {
    obj = new Controller();
    con = obj.con();
    obj.createStatement(con);

    try {
        userid = (String) session.getAttribute("UserId");
    } catch (Exception e) {
        userid = "0";
    }

    if (userid == null) {
        userid = "0";
        response.sendRedirect(request.getContextPath() + "/index.jsp");
    }

    Office_id = obj.getValu(
        "PMS_DCB_COM_OFFICE_SWITCH",
        "CASE WHEN OLD_OFFICE_ID IS NULL and DATE_ALLOWED_UPTO IS NULL THEN OFFICE_ID WHEN DATE_ALLOWED_UPTO >= clock_timestamp() and SWITCH_ID is not null THEN SWITCH_ID ELSE OFFICE_ID END AS OFFICE_ID",
        "where EMPLOYEE_ID in ( select EMPLOYEE_ID from SEC_MST_USERS where USER_ID='" + userid + "')",
        "OFFICE_ID"
    );

    if (Office_id.equalsIgnoreCase("0")) {
        Office_id = "5000";
    }

    Office_Name = obj.getValue("COM_MST_OFFICES", "OFFICE_NAME", " where OFFICE_ID=" + Office_id);

} catch (Exception e) {
    out.println("<div style='color:red;font-weight:bold;'>Main Load Error : " + e + "</div>");
}

// ===== Decide office used in all queries (global) =====
String currentOfficeId = Office_id;
if ("5000".equals(Office_id) && selectedOfficeId != null && !"0".equals(selectedOfficeId)) {
    currentOfficeId = selectedOfficeId;
}
// =====================================================

//======== OFFICE DROPDOWN VARIABLE =========
String officeDropdownOptions = "";

if (Office_id.equals("5000")) {
    PreparedStatement psOffice = null;
    ResultSet rsOffice = null;

    try {
        String sqlOffice =
            "SELECT office_id, office_name " +
            "FROM com_mst_offices where office_id in " +
            "(select office_id from pms_dcb_mst_beneficiary where status ='L' ) " +
            "ORDER BY office_name";

        psOffice = con.prepareStatement(sqlOffice);
        rsOffice = psOffice.executeQuery();

        officeDropdownOptions += "<option value='0'>- - - Select - - -</option>";

        while (rsOffice.next()) {
            String ofcId = rsOffice.getString("office_id");
            String ofcName = rsOffice.getString("office_name");
            String sel = selectedOfficeId.equals(ofcId) ? " selected='selected'" : "";
            officeDropdownOptions += "<option value='" + ofcId + "'" + sel + ">" + ofcName + "</option>";
        }
    } catch (Exception e) {
        out.println("<div style='color:red;font-weight:bold;'>Office Load Error : " + e + "</div>");
    } finally {
        try { if (rsOffice != null) rsOffice.close(); } catch (Exception ex) {}
        try { if (psOffice != null) psOffice.close(); } catch (Exception ex) {}
    }
}
//======== END OFFICE_DROPDOWN_VARIABLE =========

int prvmonth = 0;
int prvyear = 0;
try {
    prvmonth = obj.prv_month(year, month);
    prvyear = obj.prv_year(year, month);
} catch(Exception e) {
    prvmonth = month;
    prvyear = year;
}

if (selectedToYear.equals("0")) selectedToYear = String.valueOf(prvyear);
if (selectedToMonth.equals("0")) selectedToMonth = String.valueOf(prvmonth);

String[] amonth = { "-select month-", "January", "February", "March", "April", "May", "June", "July",
"August", "September", "October", "November", "December" };

String beneficiaryTypeOptions = "<option value='0'>- - - Select - - -</option>";
String beneficiaryOptions = "<option value='0'>- - - Select - - -</option>";
String districtOptions = "<option value='0'>- - - Select - - -</option>";
String blockOptions = "<option value='0'>Select Block</option>";
%>

<%
PreparedStatement psBenType = null;
ResultSet rsBenType = null;

try {
    String sqlBenType =
        "SELECT DISTINCT BEN.beneficiary_type_id_sub AS BENEFICIARY_TYPE_ID, " +
        "COALESCE(TYP.ben_type_desc, 'UNKNOWN') AS BENEFICIARY_TYPE_NAME " +
        "FROM pms_dcb_mst_beneficiary BEN " +
        "LEFT JOIN pms_dcb_ben_type TYP ON TYP.ben_type_id = BEN.beneficiary_type_id_sub " +
        "WHERE BEN.office_id = ? AND BEN.status = 'L' " +
        "ORDER BY BEN.beneficiary_type_id_sub";

    psBenType = con.prepareStatement(sqlBenType);
    psBenType.setInt(1, Integer.parseInt(currentOfficeId));
    rsBenType = psBenType.executeQuery();

    while (rsBenType.next()) {
        String benTypeId = rsBenType.getString("BENEFICIARY_TYPE_ID");
        String benTypeName = rsBenType.getString("BENEFICIARY_TYPE_NAME");

        beneficiaryTypeOptions += "<option value='" + benTypeId + "' "
            + (benTypeId.equals(selectedBeneficiaryType) ? "selected='selected'" : "")
            + ">" + benTypeName + "</option>";
    }
} catch (Exception e) {
    out.println("<div style='color:red;font-weight:bold;'>Beneficiary Type Load Error : " + e + "</div>");
    beneficiaryTypeOptions = "<option value='0'>No Data Found</option>";
} finally {
    try { if (rsBenType != null) rsBenType.close(); } catch (Exception ex) {}
    try { if (psBenType != null) psBenType.close(); } catch (Exception ex) {}
}
%>

<%
PreparedStatement psDistrict = null;
ResultSet rsDistrict = null;

try {
    String sqlDistrict =
        "SELECT district_code, district_name " +
        "FROM com_mst_districts where district_code in "+
        "(select distinct district_code from pms_dcb_mst_beneficiary where office_id = ? )" +
        "ORDER BY district_name";

    psDistrict = con.prepareStatement(sqlDistrict);
    psDistrict.setInt(1, Integer.parseInt(currentOfficeId));

    rsDistrict = psDistrict.executeQuery();

    while (rsDistrict.next()) {
        String distId = rsDistrict.getString("district_code");
        String distName = rsDistrict.getString("district_name");

        districtOptions += "<option value='" + distId + "' "
            + (distId.equals(selectedDistrictId) ? "selected='selected'" : "")
            + ">" + distName + "</option>";
    }
} catch (Exception e) {
    out.println("<div style='color:red;font-weight:bold;'>District Load Error : " + e + "</div>");
    districtOptions = "<option value='0'>No Data Found</option>";
} finally {
    try { if (rsDistrict != null) rsDistrict.close(); } catch (Exception ex) {}
    try { if (psDistrict != null) psDistrict.close(); } catch (Exception ex) {}
}
%>

<%
PreparedStatement psBlock = null;
ResultSet rsBlock = null;

try {
    if (!"0".equals(selectedDistrictId)) {
        String sqlBlock =
            "SELECT block_sno, block_name " +
            "FROM com_mst_blocks " +
            "WHERE district_code = ?::numeric " +
            "ORDER BY block_name";

        psBlock = con.prepareStatement(sqlBlock);
        psBlock.setString(1, selectedDistrictId);
        rsBlock = psBlock.executeQuery();

        StringBuffer blk = new StringBuffer();
        blk.append("<option value='0'>Select Block</option>");

        while (rsBlock.next()) {
            String blockCode = rsBlock.getString("block_sno");
            String blockName = rsBlock.getString("block_name");

            blk.append("<option value='").append(blockCode).append("' ");
            if (blockCode.equals(selectedBlockId)) {
                blk.append("selected='selected' ");
            }
            blk.append(">").append(blockName).append("</option>");
        }

        blockOptions = blk.toString();
    }
} catch (Exception e) {
    out.println("<div style='color:red;font-weight:bold;'>Block Load Error : " + e + "</div>");
    blockOptions = "<option value='0'>No Data Found</option>";
} finally {
    try { if (rsBlock != null) rsBlock.close(); } catch (Exception ex) {}
    try { if (psBlock != null) psBlock.close(); } catch (Exception ex) {}
}
%>

<%
PreparedStatement psBeneficiary = null;
ResultSet rsBeneficiary = null;

try {
    if (!selectedBeneficiaryType.equals("0")) {

        String sqlBeneficiary = "";

        if ("6".equals(selectedBeneficiaryType)) {
            sqlBeneficiary =
                "SELECT beneficiary_sno, beneficiary_name " +
                "FROM pms_dcb_mst_beneficiary " +
                "WHERE beneficiary_type_id_sub = ? " +
                "AND office_id = ? " +
                "AND status = 'L' " +
                "AND district_code = ?::numeric " +
                "AND block_sno = ?::numeric " +
                "ORDER BY beneficiary_name";
            
            out.println("<!-- DEBUG: type=" + selectedBeneficiaryType +
                    ", office=" + currentOfficeId +
                    ", dist=" + selectedDistrictId +
                    ", block=" + selectedBlockId + " -->");

            psBeneficiary = con.prepareStatement(sqlBeneficiary);
            psBeneficiary.setInt(1, Integer.parseInt(selectedBeneficiaryType));
            psBeneficiary.setInt(2, Integer.parseInt(currentOfficeId));
            psBeneficiary.setString(3, selectedDistrictId);
            psBeneficiary.setString(4, selectedBlockId);

        } else {
            sqlBeneficiary =
                "SELECT beneficiary_sno, beneficiary_name " +
                "FROM pms_dcb_mst_beneficiary " +
                "WHERE beneficiary_type_id_sub = ? " +
                "AND office_id = ? " +
                "AND status = 'L' " +
                "ORDER BY beneficiary_name";

            psBeneficiary = con.prepareStatement(sqlBeneficiary);
            psBeneficiary.setInt(1, Integer.parseInt(selectedBeneficiaryType));
            psBeneficiary.setInt(2, Integer.parseInt(currentOfficeId));
        }

        rsBeneficiary = psBeneficiary.executeQuery();

        beneficiaryOptions = "<option value='0'>- - - Select - - -</option>";

        while (rsBeneficiary.next()) {
            String benId = rsBeneficiary.getString("beneficiary_sno");
            String benName = rsBeneficiary.getString("beneficiary_name");

            beneficiaryOptions += "<option value='" + benId + "' "
                + (benId.equals(selectedBeneficiaryId) ? "selected='selected'" : "")
                + ">" + benName + " (" + benId + ")</option>";
        }
    }
} catch (Exception e) {
    out.println("<div style='color:red;font-weight:bold;'>Beneficiary Load Error : " + e + "</div>");
    beneficiaryOptions = "<option value='0'>No Data Found</option>";
} finally {
    try { if (rsBeneficiary != null) rsBeneficiary.close(); } catch (Exception ex) {}
    try { if (psBeneficiary != null) psBeneficiary.close(); } catch (Exception ex) {}
}
%>

<% if (!isExcel) { %>
<form name="pms_dcb_ledger_report" method="post">
<input type="hidden" name="mode" id="mode" value="" />

<table border="1" width="75%" align="center" cellpadding="4" cellspacing="0" class="alerts2">
<tr bgcolor="#408099">
<td colspan="4" align="center"><b>DCB Beneficiary Monthwise Report</b></td>
</tr>

<tr class="tdText">
<td width="20%">Division Name</td>
<td width="30%"><font color="blue"><%=Office_Name%></font></td>
</tr>

<% if ("5000".equals(Office_id)) { %>
<tr class="tdText">
    <td width="20%">Office</td>
    <td colspan="3">
        <select id="officeId" name="officeId" style="width:300px" class="select"
                onchange="onOfficeChange();">
            <%=officeDropdownOptions%>
        </select>
    </td>
</tr>
<% } %>

<tr class="tdText">
<td width="20%">From Year</td>
<td width="30%">
<select id="fromYear" name="fromYear" style="width: 160px" class="select">
<option value="0">- - - Select - - -</option>
<%
for (int j = 2009; j <= year; j++) {
%>
<option value="<%=j%>" <%= String.valueOf(j).equals(selectedFromYear) ? "selected='selected'" : "" %>><%=j%></option>
<%
}
%>
</select>
</td>

<td width="20%">From Month</td>
<td width="30%">
<select id="fromMonth" name="fromMonth" style="width: 160px" class="select">
<option value="0">- - - Select - - -</option>
<%
for (int i = 1; i <= 12; i++) {
%>
<option value="<%=i%>" <%= String.valueOf(i).equals(selectedFromMonth) ? "selected='selected'" : "" %>><%=amonth[i]%></option>
<%
}
%>
</select>
</td>
</tr>

<tr class="tdText">
<td width="20%">To Year</td>
<td width="30%">
<select id="toYear" name="toYear" style="width: 160px" class="select">
<option value="0">- - - Select - - -</option>
<%
for (int j = year - 6; j <= year; j++) {
%>
<option value="<%=j%>" <%= String.valueOf(j).equals(selectedToYear) ? "selected='selected'" : "" %>><%=j%></option>
<%
}
%>
</select>
</td>

<td width="20%">To Month</td>
<td width="30%">
<select id="toMonth" name="toMonth" style="width: 160px" class="select">
<option value="0">- - - Select - - -</option>
<%
for (int i = 1; i <= 12; i++) {
%>
<option value="<%=i%>" <%= String.valueOf(i).equals(selectedToMonth) ? "selected='selected'" : "" %>><%=amonth[i]%></option>
<%
}
%>
</select>
</td>
</tr>

<tr class="tdText">
<td width="20%">Beneficiary Type</td>
<td colspan="3">
<select id="beneficiaryType" name="beneficiaryType" style="width:220px" class="select"
        onchange="onBeneficiaryTypeChange();">
<%=beneficiaryTypeOptions%>
</select>
</td>
</tr>

<tr class="tdText" id="districtRow" style="<%= "6".equals(selectedBeneficiaryType) ? "" : "display:none;" %>">
<td width="20%">District</td>
<td colspan="3">
<select id="districtId" name="districtId" style="width:220px" class="select"
        onchange="getBlocks(this.value);">
<%=districtOptions%>
</select>
</td>
</tr>

<tr class="tdText" id="blockRow" style="<%= "6".equals(selectedBeneficiaryType) ? "" : "display:none;" %>">
<td width="20%">Block</td>
<td colspan="3">
<select id="blockId" name="blockId" style="width:220px" class="select"
        onchange="this.form.submit();">
<%=blockOptions%>
</select>
</td>
</tr>

<tr class="tdText">
<td width="20%">Beneficiary</td>
<td colspan="3">
<select id="beneficiaryId" name="beneficiaryId" style="width:220px" class="select">
<%=beneficiaryOptions%>
</select>
</td>
</tr>

<tr>
<td colspan="4" align="center">
<input type="button"
name="Print"
value="View"
class="bprint"
id="cmdprint"
onclick="funcmddata();"
style="font-style: italic; font-weight: bold; font-size: 7; color: red" />

<input type="button"
       name="Excel"
       value="Export Excel"
       class="bprint"
       onclick="exportExcel();"
       style="font-style: italic; font-weight: bold; font-size: 7; color: green" />
&nbsp;

<input type="button"
name="Clear"
value="Clear"
id="clear"
onclick="refreshForm();"
style="font-style: italic; font-weight: bold; font-size: 7; color: blue"
class="fb2" />

<input type="button"
name="Exit"
value="Exit"
id="exit"
onclick="exitwindow();"
style="font-style: italic; font-weight: bold; font-size: 7; color: red"
class="fb2" />


</td>
</tr>
</table>
</form>
<% } %>

<%
if ("print".equalsIgnoreCase(mode)) {
    PreparedStatement psReport = null;
    ResultSet rsReport = null;

    try {
    	String sqlReport =
    			"SELECT bill_year as year , initcap(get_month_name(bill_month)) as month, bill_sno,dmd.beneficiary_sno ,beneficiary_type,district_name ," +
    			"block_name , panchayat_name , beneficiary_name , " +
    			"COALESCE(qty.qty_consumed, 0) as quantity_consumed, " +
    			"outstanding_due_wc , collection_wc , demand_wc , other_charges_wc , balance_wc ," +
    			"outstanding_due_int , collection_int , demand_int , other_charges_int , balance_int ," +
    			"outstanding_due_maint , collection_maint , other_charges_maint , balance_maint ," +
    			"outstanding_due , collection, demand , other_charges , balance " +
    			"FROM full_view_new dmd " +
    			"LEFT JOIN ( " +
    			"  SELECT beneficiary_sno , year , month , SUM(qty_consumed) AS qty_consumed " +
    			"  FROM pms_dcb_wc_billing " +
    			"  GROUP BY beneficiary_sno , year , month " +
    			") qty " +
    			"ON qty.beneficiary_sno = dmd.beneficiary_sno " +
    			"AND qty.year = bill_year " +
    			"AND qty.month = bill_month " +
    			"WHERE dmd.beneficiary_sno = ? " +
    			"AND (bill_year, bill_month) BETWEEN (?, ?) AND (?, ?) " +
    			"ORDER BY bill_year, bill_month";

        psReport = con.prepareStatement(sqlReport);
        psReport.setInt(1, Integer.parseInt(selectedBeneficiaryId));
        psReport.setInt(2, Integer.parseInt(selectedFromYear));
        psReport.setInt(3, Integer.parseInt(selectedFromMonth));
        psReport.setInt(4, Integer.parseInt(selectedToYear));
        psReport.setInt(5, Integer.parseInt(selectedToMonth));

        rsReport = psReport.executeQuery();
        ResultSetMetaData rsmd = rsReport.getMetaData();
        int colCount = rsmd.getColumnCount();
        boolean hasData = false;
%>
<br>

<div class="reportScroll">
<table border="1" cellpadding="4" cellspacing="0" class="alerts2">
<tr>
<td colspan="<%=colCount + 1%>" align="center" style="background-color:#408099; color:white;">
<b>Beneficiary Monthwise Report</b>
</td>
</tr>

<tr class="tdText">
<th>S.No</th>
<%
for (int c = 1; c <= colCount; c++) {
%>
<th><%=formatColumnName(rsmd.getColumnLabel(c))%></th>
<%
}
%>
</tr>

<%
int rownum = 1;
while (rsReport.next()) {
hasData = true;
%>
<tr class="tdText">
<td><%=rownum++%></td>
<%
for (int c = 1; c <= colCount; c++) {
%>
<td><%=rsReport.getString(c) == null ? "" : rsReport.getString(c)%></td>
<%
}
%>
</tr>
<%
}

if (!hasData) {
%>
<tr class="tdText">
<td colspan="<%=colCount + 1%>" align="center">
<font color="red"><b>No Records Found</b></font>
</td>
</tr>
<%
}
%>
</table>
</div>

<%
    } catch (Exception e) {
        out.println("<div style='width:95%;margin:10px auto;color:red;font-weight:bold;'>Report Load Error : " + e + "</div>");
    } finally {
        try { if (rsReport != null) rsReport.close(); } catch (Exception e) {}
        try { if (psReport != null) psReport.close(); } catch (Exception e) {}
    }
}

try { if (con != null) con.close(); } catch (Exception e) {}
%>

</body>
</html>