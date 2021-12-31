<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="DXWebApplication1.WebForm1" %>

<%@ Register Assembly="DevExpress.Web.ASPxSpreadsheet.v18.2, Version=18.2.16.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxSpreadsheet" TagPrefix="dx" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script>
        function OnCommandExecuted(s, e) {
            var command = e.item.name;
            ASPxSpreadsheet1.PerformCallback(command)
        }

    </script>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <dx:ASPxRibbon runat="server" ID="ASPxRibbon1" ShowFileTab="false" ShowTabs="false" OneLineMode="true">
                <ClientSideEvents CommandExecuted="OnCommandExecuted" />
                <Tabs>
                    <dx:RibbonTab>
                        <Groups>
                            <dx:RibbonGroup  >
                                <Items>
                                    <dx:RibbonButtonItem Text="Apply formatting" Name="applyFormatting"></dx:RibbonButtonItem>
                                    <dx:RibbonButtonItem Text="Insert link" Name="insertLink"></dx:RibbonButtonItem>
                                    <dx:RibbonButtonItem Text="Draw Borders" Name="drawBorders"></dx:RibbonButtonItem>
                                    <dx:RibbonButtonItem Text="Show total" Name="showTotal"></dx:RibbonButtonItem>
                                </Items>
                            </dx:RibbonGroup>
                        </Groups>
                    </dx:RibbonTab>
                </Tabs>
            </dx:ASPxRibbon>
            <dx:ASPxSpreadsheet runat="server" ID="ASPxSpreadsheet1" ClientInstanceName="ASPxSpreadsheet1" OnCallback="ASPxSpreadsheet1_Callback"></dx:ASPxSpreadsheet>
        </div>
    </form>
</body>
</html>
