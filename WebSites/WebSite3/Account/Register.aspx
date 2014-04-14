<%@ Page Title="Регистрация" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeFile="Register.aspx.cs" Inherits="Account_Register" %>

<asp:Content runat="server" ID="BodyContent" ContentPlaceHolderID="MainContent">
    <h2><%: Title %>.</h2>
    <p class="text-danger">
        <asp:Literal runat="server" ID="ErrorMessage" />
    </p>

    <div class="form-horizontal">
        <h4>Создание новой учетной записи.</h4>
        <hr />
        <asp:ValidationSummary runat="server" CssClass="text-danger" />
        <div class="form-group">
            <asp:Label runat="server" AssociatedControlID="UserName" CssClass="col-md-2 control-label">Имя пользователя</asp:Label>
            <div class="col-md-10">
                <asp:TextBox runat="server" ID="UserName" CssClass="form-control" />
                <asp:RequiredFieldValidator runat="server" ControlToValidate="UserName"
                    CssClass="text-danger" ErrorMessage="Поле имени пользователя заполнять обязательно." />
            </div>
        </div>
        <div class="form-group">
            <asp:Label runat="server" AssociatedControlID="Password" CssClass="col-md-2 control-label">Пароль</asp:Label>
            <div class="col-md-10">
                <asp:TextBox runat="server" ID="Password" TextMode="Password" CssClass="form-control" />
                <asp:RequiredFieldValidator runat="server" ControlToValidate="Password"
                    CssClass="text-danger" ErrorMessage="Поле пароля заполнять обязательно." />
            </div>
        </div>
        <div class="form-group">
            <asp:Label runat="server" AssociatedControlID="ConfirmPassword" CssClass="col-md-2 control-label">Подтверждение пароля</asp:Label>
            <div class="col-md-10">
                <asp:TextBox runat="server" ID="ConfirmPassword" TextMode="Password" CssClass="form-control" />
                <asp:RequiredFieldValidator runat="server" ControlToValidate="ConfirmPassword"
                    CssClass="text-danger" Display="Dynamic" ErrorMessage="Поле подтверждения пароля заполнять обязательно." />
                <asp:CompareValidator runat="server" ControlToCompare="Password" ControlToValidate="ConfirmPassword"
                    CssClass="text-danger" Display="Dynamic" ErrorMessage="Пароль и его подтверждение не совпадают." />
            </div>
        </div>
        <div class="form-group">
            <div class="col-md-offset-2 col-md-10">
                <br />
                <asp:TextBox ID="FamilyNameTxtBox" runat="server"></asp:TextBox>
                <asp:Label ID="Label1" runat="server" Text="Имя и Фамилия"></asp:Label>
                <br />
                <asp:TextBox ID="GroupNameTxtBox" runat="server"></asp:TextBox>
                <asp:Label ID="Label2" runat="server" Text="Номер группы полностью"></asp:Label>
                <br />
                <asp:TextBox ID="PhNumberTxtBox" runat="server"></asp:TextBox>
                <asp:Label ID="Label3" runat="server" Text="Номер телефона"></asp:Label>
                <br />
                <asp:TextBox ID="MailTxtBox" runat="server"></asp:TextBox>
                <asp:Label ID="Label4" runat="server" Text="E-mail"></asp:Label>
                <br />
                <asp:TextBox ID="SweetTxtBox" runat="server"></asp:TextBox>
                <asp:Label ID="Label5" runat="server" Text="День Варенья"></asp:Label>
                <br />
                <asp:Button ID="Button1" runat="server" OnClick="CreateUser_Click" Text="Регистрация" CssClass="btn btn-default" />
            </div>
        </div>
    </div>
</asp:Content>

