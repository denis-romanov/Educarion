<%@ Page Title="" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true" CodeFile="Rasp.aspx.cs" Inherits="Rasp" %>

<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" Runat="Server">
    

        <asp:Label ID="Label1" runat="server" Text="Введите имя таблицы полностью"></asp:Label>
        <br />
        <asp:TextBox ID="TextBox1" runat="server" BackColor="White" BorderColor="Black" BorderStyle="Solid"></asp:TextBox>
        <br />
    

        Введите имя группы полностью<br />
    

        <asp:TextBox ID="TextBoxNgrup" runat="server" BorderColor="Black" BorderStyle="Solid"></asp:TextBox>

    
        <br />
        <asp:Button ID="ButtonShowRasp" runat="server" OnClick="ButtonShowRasp_Click" Text="Показать расписание" />
        <br />
    

    <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" DataSourceID="SqlDataSource1">
        <Columns>
            <asp:BoundField DataField="Дни" HeaderText="Дни" SortExpression="Дни" />
            <asp:BoundField DataField="Часы" HeaderText="Часы" SortExpression="Часы" />
            <asp:BoundField DataField="Неделя" HeaderText="Неделя" SortExpression="Неделя" />
            <asp:BoundField DataField="3ПСУ-1ДБ-228" HeaderText="3ПСУ-1ДБ-228" SortExpression="3ПСУ-1ДБ-228" />
        </Columns>
    </asp:GridView>
    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:TestConnectionString %>" SelectCommand="RaspShow" SelectCommandType="StoredProcedure">
        <SelectParameters>
            <asp:ControlParameter ControlID="TextBoxNgrup" Name="col_name" PropertyName="Text" Type="String" />
            <asp:ControlParameter ControlID="TextBox1" Name="tab_name" PropertyName="Text" Type="String" />
        </SelectParameters>
    </asp:SqlDataSource>
</asp:Content>

