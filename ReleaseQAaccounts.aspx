<%@ Page Title="" Language="C#" MasterPageFile="~/ARMS.master" AutoEventWireup="true" CodeFile="ReleaseQAaccounts.aspx.cs" Inherits="WorkAllocation_ReleaseQAaccounts" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <link href="../css/iArmsStyleSheet1.css" rel="stylesheet" type="text/css" />
    <link href="../DataTables/DataTables-1.10.13/css/jquery.dataTables.css" rel="stylesheet" />
    <link href="../DataTables/Bootstrap-3.3.7/css/bootstrap.css" rel="stylesheet" />
    <link href="../CSS/sumoselect.css" rel="stylesheet" />
    <link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css" />
    <script type="text/javascript" src="../DataTables/jQuery-2.2.4/jquery-2.2.4.js"></script>
    <script type="text/javascript" src="../DataTables/DataTables-1.10.13/js/jquery.dataTables.js"></script>
    <script type="text/javascript" src="../Scripts/jquery.sumoselect.js"></script>
    <script type="text/javascript" src="../Scripts/jquery-ui.js"></script>
    <script type="text/javascript" src="../Scripts/bootstrap.min.js"></script>

    <style>
        .innerStyle {
            font-family: verdana;
            font-size: 11px;
            background-color: White;
            Width: 100%;
        }

        table.dataTable tr {
            background-color: White;
            Height: 10px;
            font-family: Verdana;
            font-size: 11px;
            color: black;
        }

            table.dataTable tr:nth-child(even) {
                background-color: #E6F2FF;
                height: 10px;
                font-family: Verdana;
                font-size: 11px;
                color: black;
            }

        table.dataTable tfoot {
            display: table-header-group;
        }

        tbody th, table.dataTable tbody td {
            padding: 0px 0px;
            border: 1px solid #7b7b7b;
            border-collapse: collapse;
            border-width: thin;
            text-align: center;
            height: 23px;
        }

        #grdReleaseQAaccounts th {
            top: expression(document.getElementById("ctl00_cphPlaceHolder_pnlAccountDetails1").scrollTop-1);
            left: expression(parentNode.parentNode.parentNode.parentNode.scrollLeft);
            position: static;
            z-index: 0;
            font-family: Verdana, Geneva, sans-serif;
            background-color: #4C72AA;
            background-repeat: repeat-x;
            height: 20px;
            color: White;
            text-decoration: none;
            width: 30px;
        }

        div#grdReleaseQAaccounts th {
            top: expression(document.getElementById("cphPlaceHolder_pnlAccountDetails1").scrollTop-1);
            left: expression(parentNode.parentNode.parentNode.parentNode.scrollLeft);
            position: static;
            z-index: 0;
            font-family: Verdana, Geneva, sans-serif;
            background-color: #4C72AA;
            background-repeat: repeat-x;
            height: 20px;
            color: White;
            text-decoration: none;
            width: 30px;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="cphPlaceHolder" runat="Server">

    <div class="col-lg-12" style="padding: 10px;">
        <div class="col-lg-12" style="padding: 10px; border: 1px solid #eeeeee">
            <input type="button" value="Unlock" id="btnUnlock" style="left:10px" class="btn btn-primary" />
        </div>
        <div class="col-lg-12" style="padding: 10px; border: 1px solid #eeeeee">
            <div class="col-lg-1"></div>
            <div class="col-lg-12">
                <div class="progress" id="progress">
                    <img alt="imgload" class="imga" src="../Images/loading.gif" />
                </div>
                <table id="grdReleaseQAaccounts" class="innerStyle dataTable" style="width: 100%;">

                    <thead>
                        <tr>
                            <th>Sl No</th>
                            <th>
                                <input type="checkbox" />
                             </th>
                            <th>Locked By</th>
                            <th>Claim No</th>
                            <th>Status</th>
                            <th>Patient Acct</th>
                            <th>Status Date</th>
                            <th>Created Date</th>
                            <th>Service Date</th>
                            <th>Payer</th>
                            <th>Provider Name</th>
                            <th>Charges</th>
                        </tr>
                    </thead>


                </table>
            </div>
            <div class="col-lg-1" ></div>
        </div>
    </div>


    <script type="text/javascript">
        $(document).ready(function () {
            BindGrid();
            $('#btnUnlock').click(function (e) {
                var selectedProcId = '';
                if ($('#grdReleaseQAaccounts tbody').children().length > 0) {
                    $('#grdReleaseQAaccounts tr td').find('input[type=checkbox]').each(function (index, element) {
                        if (element.checked == true) {

                            if (selectedProcId != '')
                                selectedProcId = selectedProcId + ',' + element.value;
                            else
                                selectedProcId = element.value;
                        }

                        
                    });

                }

                console.log(selectedProcId);
                if(selectedProcId!="")
                {
                    alert(selectedProcId);
                }
            });
        });
        function BindGrid()
        {
            $.ajax({
                beforeSend: function () {
                    $("#progress").show();
                },
                type: "POST",
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                url: "ReleaseQAaccounts.aspx/GetLockedAccounts",
                success: function (data) {
                    var datatableVariable = $('#grdReleaseQAaccounts').DataTable({
                        data: data.d,
                        columns: [
                            { 'data': 'SlNo', 'name': 'Sl No' },
                            { 'data': 'ProcId', 'render': function (ProcId) {
                                    return '<input type="checkbox" value=' + ProcId + ' />';
                                }, 'sorting': 'false',
                            },
                            { 'data': 'Locked_by' },
                            { 'data': 'Claim_No' },
                            { 'data': 'Status' },
                            { 'data': 'Patient_Acct' },
                            { 'data': 'Status_Date' },
                            { 'data': 'Created_Date' },
                            { 'data': 'Service_Date' },
                            { 'data': 'Payer' },
                            { 'data': 'Provider_Name' },
                            { 'data': 'Charges' }

                        ]
                    });

                }, complete: function () {
                    $("#progress").hide();

                }
            });
        }
    </script>
</asp:Content>

