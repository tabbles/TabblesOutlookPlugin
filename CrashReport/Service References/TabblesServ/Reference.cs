﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace CrashReport.TabblesServ {
    using System.Runtime.Serialization;
    using System;
    
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Runtime.Serialization", "4.0.0.0")]
    [System.Runtime.Serialization.DataContractAttribute(Name="Tabbles4NewVersionData", Namespace="http://schemas.datacontract.org/2004/07/UpdatesWebService")]
    [System.SerializableAttribute()]
    public partial class Tabbles4NewVersionData : object, System.Runtime.Serialization.IExtensibleDataObject, System.ComponentModel.INotifyPropertyChanged {
        
        [System.NonSerializedAttribute()]
        private System.Runtime.Serialization.ExtensionDataObject extensionDataField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private System.Nullable<int> betaNumberField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private System.Nullable<System.DateTime> buildDateOfLatestVersionField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string changeLogField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string customTextField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string latestVersionField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string urlCompleteChangelogField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string urlToDownloadLatestField;
        
        [global::System.ComponentModel.BrowsableAttribute(false)]
        public System.Runtime.Serialization.ExtensionDataObject ExtensionData {
            get {
                return this.extensionDataField;
            }
            set {
                this.extensionDataField = value;
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public System.Nullable<int> betaNumber {
            get {
                return this.betaNumberField;
            }
            set {
                if ((this.betaNumberField.Equals(value) != true)) {
                    this.betaNumberField = value;
                    this.RaisePropertyChanged("betaNumber");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public System.Nullable<System.DateTime> buildDateOfLatestVersion {
            get {
                return this.buildDateOfLatestVersionField;
            }
            set {
                if ((this.buildDateOfLatestVersionField.Equals(value) != true)) {
                    this.buildDateOfLatestVersionField = value;
                    this.RaisePropertyChanged("buildDateOfLatestVersion");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string changeLog {
            get {
                return this.changeLogField;
            }
            set {
                if ((object.ReferenceEquals(this.changeLogField, value) != true)) {
                    this.changeLogField = value;
                    this.RaisePropertyChanged("changeLog");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string customText {
            get {
                return this.customTextField;
            }
            set {
                if ((object.ReferenceEquals(this.customTextField, value) != true)) {
                    this.customTextField = value;
                    this.RaisePropertyChanged("customText");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string latestVersion {
            get {
                return this.latestVersionField;
            }
            set {
                if ((object.ReferenceEquals(this.latestVersionField, value) != true)) {
                    this.latestVersionField = value;
                    this.RaisePropertyChanged("latestVersion");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string urlCompleteChangelog {
            get {
                return this.urlCompleteChangelogField;
            }
            set {
                if ((object.ReferenceEquals(this.urlCompleteChangelogField, value) != true)) {
                    this.urlCompleteChangelogField = value;
                    this.RaisePropertyChanged("urlCompleteChangelog");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string urlToDownloadLatest {
            get {
                return this.urlToDownloadLatestField;
            }
            set {
                if ((object.ReferenceEquals(this.urlToDownloadLatestField, value) != true)) {
                    this.urlToDownloadLatestField = value;
                    this.RaisePropertyChanged("urlToDownloadLatest");
                }
            }
        }
        
        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
        
        protected void RaisePropertyChanged(string propertyName) {
            System.ComponentModel.PropertyChangedEventHandler propertyChanged = this.PropertyChanged;
            if ((propertyChanged != null)) {
                propertyChanged(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
            }
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Runtime.Serialization", "4.0.0.0")]
    [System.Runtime.Serialization.DataContractAttribute(Name="Tabbles4CloudSqlServerData", Namespace="http://schemas.datacontract.org/2004/07/UpdatesWebService")]
    [System.SerializableAttribute()]
    public partial class Tabbles4CloudSqlServerData : object, System.Runtime.Serialization.IExtensibleDataObject, System.ComponentModel.INotifyPropertyChanged {
        
        [System.NonSerializedAttribute()]
        private System.Runtime.Serialization.ExtensionDataObject extensionDataField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string SqlLoginNameField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string SqlLoginPasswordField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string SqlServerAddressField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string SqlServerDatabaseNameField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string SqlServerPortField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string UrlForgottenPasswordCloudField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string UrlRegisterCloudField;
        
        [global::System.ComponentModel.BrowsableAttribute(false)]
        public System.Runtime.Serialization.ExtensionDataObject ExtensionData {
            get {
                return this.extensionDataField;
            }
            set {
                this.extensionDataField = value;
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string SqlLoginName {
            get {
                return this.SqlLoginNameField;
            }
            set {
                if ((object.ReferenceEquals(this.SqlLoginNameField, value) != true)) {
                    this.SqlLoginNameField = value;
                    this.RaisePropertyChanged("SqlLoginName");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string SqlLoginPassword {
            get {
                return this.SqlLoginPasswordField;
            }
            set {
                if ((object.ReferenceEquals(this.SqlLoginPasswordField, value) != true)) {
                    this.SqlLoginPasswordField = value;
                    this.RaisePropertyChanged("SqlLoginPassword");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string SqlServerAddress {
            get {
                return this.SqlServerAddressField;
            }
            set {
                if ((object.ReferenceEquals(this.SqlServerAddressField, value) != true)) {
                    this.SqlServerAddressField = value;
                    this.RaisePropertyChanged("SqlServerAddress");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string SqlServerDatabaseName {
            get {
                return this.SqlServerDatabaseNameField;
            }
            set {
                if ((object.ReferenceEquals(this.SqlServerDatabaseNameField, value) != true)) {
                    this.SqlServerDatabaseNameField = value;
                    this.RaisePropertyChanged("SqlServerDatabaseName");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string SqlServerPort {
            get {
                return this.SqlServerPortField;
            }
            set {
                if ((object.ReferenceEquals(this.SqlServerPortField, value) != true)) {
                    this.SqlServerPortField = value;
                    this.RaisePropertyChanged("SqlServerPort");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string UrlForgottenPasswordCloud {
            get {
                return this.UrlForgottenPasswordCloudField;
            }
            set {
                if ((object.ReferenceEquals(this.UrlForgottenPasswordCloudField, value) != true)) {
                    this.UrlForgottenPasswordCloudField = value;
                    this.RaisePropertyChanged("UrlForgottenPasswordCloud");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string UrlRegisterCloud {
            get {
                return this.UrlRegisterCloudField;
            }
            set {
                if ((object.ReferenceEquals(this.UrlRegisterCloudField, value) != true)) {
                    this.UrlRegisterCloudField = value;
                    this.RaisePropertyChanged("UrlRegisterCloud");
                }
            }
        }
        
        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
        
        protected void RaisePropertyChanged(string propertyName) {
            System.ComponentModel.PropertyChangedEventHandler propertyChanged = this.PropertyChanged;
            if ((propertyChanged != null)) {
                propertyChanged(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
            }
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Runtime.Serialization", "4.0.0.0")]
    [System.Runtime.Serialization.DataContractAttribute(Name="serverData", Namespace="http://schemas.datacontract.org/2004/07/UpdatesWebService")]
    [System.SerializableAttribute()]
    public partial class serverData : object, System.Runtime.Serialization.IExtensibleDataObject, System.ComponentModel.INotifyPropertyChanged {
        
        [System.NonSerializedAttribute()]
        private System.Runtime.Serialization.ExtensionDataObject extensionDataField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string passwordField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string serverField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string userNameField;
        
        [global::System.ComponentModel.BrowsableAttribute(false)]
        public System.Runtime.Serialization.ExtensionDataObject ExtensionData {
            get {
                return this.extensionDataField;
            }
            set {
                this.extensionDataField = value;
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string password {
            get {
                return this.passwordField;
            }
            set {
                if ((object.ReferenceEquals(this.passwordField, value) != true)) {
                    this.passwordField = value;
                    this.RaisePropertyChanged("password");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string server {
            get {
                return this.serverField;
            }
            set {
                if ((object.ReferenceEquals(this.serverField, value) != true)) {
                    this.serverField = value;
                    this.RaisePropertyChanged("server");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string userName {
            get {
                return this.userNameField;
            }
            set {
                if ((object.ReferenceEquals(this.userNameField, value) != true)) {
                    this.userNameField = value;
                    this.RaisePropertyChanged("userName");
                }
            }
        }
        
        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
        
        protected void RaisePropertyChanged(string propertyName) {
            System.ComponentModel.PropertyChangedEventHandler propertyChanged = this.PropertyChanged;
            if ((propertyChanged != null)) {
                propertyChanged(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
            }
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Runtime.Serialization", "4.0.0.0")]
    [System.Runtime.Serialization.DataContractAttribute(Name="NewVersionData", Namespace="http://schemas.datacontract.org/2004/07/UpdatesWebService")]
    [System.SerializableAttribute()]
    public partial class NewVersionData : object, System.Runtime.Serialization.IExtensibleDataObject, System.ComponentModel.INotifyPropertyChanged {
        
        [System.NonSerializedAttribute()]
        private System.Runtime.Serialization.ExtensionDataObject extensionDataField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private System.Nullable<int> betaNumberField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private System.Nullable<System.DateTime> buildDateOfLatestVersionField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string changeLogField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string customTextField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string latestVersionField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string urlCompleteChangelogField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string urlToDownloadLatest32bitField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string urlToDownloadLatest64bitField;
        
        [System.Runtime.Serialization.OptionalFieldAttribute()]
        private string urlToDownloadLatestPortableField;
        
        [global::System.ComponentModel.BrowsableAttribute(false)]
        public System.Runtime.Serialization.ExtensionDataObject ExtensionData {
            get {
                return this.extensionDataField;
            }
            set {
                this.extensionDataField = value;
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public System.Nullable<int> betaNumber {
            get {
                return this.betaNumberField;
            }
            set {
                if ((this.betaNumberField.Equals(value) != true)) {
                    this.betaNumberField = value;
                    this.RaisePropertyChanged("betaNumber");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public System.Nullable<System.DateTime> buildDateOfLatestVersion {
            get {
                return this.buildDateOfLatestVersionField;
            }
            set {
                if ((this.buildDateOfLatestVersionField.Equals(value) != true)) {
                    this.buildDateOfLatestVersionField = value;
                    this.RaisePropertyChanged("buildDateOfLatestVersion");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string changeLog {
            get {
                return this.changeLogField;
            }
            set {
                if ((object.ReferenceEquals(this.changeLogField, value) != true)) {
                    this.changeLogField = value;
                    this.RaisePropertyChanged("changeLog");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string customText {
            get {
                return this.customTextField;
            }
            set {
                if ((object.ReferenceEquals(this.customTextField, value) != true)) {
                    this.customTextField = value;
                    this.RaisePropertyChanged("customText");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string latestVersion {
            get {
                return this.latestVersionField;
            }
            set {
                if ((object.ReferenceEquals(this.latestVersionField, value) != true)) {
                    this.latestVersionField = value;
                    this.RaisePropertyChanged("latestVersion");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string urlCompleteChangelog {
            get {
                return this.urlCompleteChangelogField;
            }
            set {
                if ((object.ReferenceEquals(this.urlCompleteChangelogField, value) != true)) {
                    this.urlCompleteChangelogField = value;
                    this.RaisePropertyChanged("urlCompleteChangelog");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string urlToDownloadLatest32bit {
            get {
                return this.urlToDownloadLatest32bitField;
            }
            set {
                if ((object.ReferenceEquals(this.urlToDownloadLatest32bitField, value) != true)) {
                    this.urlToDownloadLatest32bitField = value;
                    this.RaisePropertyChanged("urlToDownloadLatest32bit");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string urlToDownloadLatest64bit {
            get {
                return this.urlToDownloadLatest64bitField;
            }
            set {
                if ((object.ReferenceEquals(this.urlToDownloadLatest64bitField, value) != true)) {
                    this.urlToDownloadLatest64bitField = value;
                    this.RaisePropertyChanged("urlToDownloadLatest64bit");
                }
            }
        }
        
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string urlToDownloadLatestPortable {
            get {
                return this.urlToDownloadLatestPortableField;
            }
            set {
                if ((object.ReferenceEquals(this.urlToDownloadLatestPortableField, value) != true)) {
                    this.urlToDownloadLatestPortableField = value;
                    this.RaisePropertyChanged("urlToDownloadLatestPortable");
                }
            }
        }
        
        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
        
        protected void RaisePropertyChanged(string propertyName) {
            System.ComponentModel.PropertyChangedEventHandler propertyChanged = this.PropertyChanged;
            if ((propertyChanged != null)) {
                propertyChanged(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
            }
        }
    }
    
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
    [System.ServiceModel.ServiceContractAttribute(ConfigurationName="TabblesServ.ITabblesService")]
    public interface ITabblesService {
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/ITabblesService/GetTabbles4NewVersionData", ReplyAction="http://tempuri.org/ITabblesService/GetTabbles4NewVersionDataResponse")]
        CrashReport.TabblesServ.Tabbles4NewVersionData GetTabbles4NewVersionData();
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/ITabblesService/GetTabbles4NewVersionData", ReplyAction="http://tempuri.org/ITabblesService/GetTabbles4NewVersionDataResponse")]
        System.Threading.Tasks.Task<CrashReport.TabblesServ.Tabbles4NewVersionData> GetTabbles4NewVersionDataAsync();
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/ITabblesService/GetTabbles4CloudSqlServerData", ReplyAction="http://tempuri.org/ITabblesService/GetTabbles4CloudSqlServerDataResponse")]
        CrashReport.TabblesServ.Tabbles4CloudSqlServerData GetTabbles4CloudSqlServerData();
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/ITabblesService/GetTabbles4CloudSqlServerData", ReplyAction="http://tempuri.org/ITabblesService/GetTabbles4CloudSqlServerDataResponse")]
        System.Threading.Tasks.Task<CrashReport.TabblesServ.Tabbles4CloudSqlServerData> GetTabbles4CloudSqlServerDataAsync();
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/ITabblesService/GetYellowBlueSoftServer", ReplyAction="http://tempuri.org/ITabblesService/GetYellowBlueSoftServerResponse")]
        CrashReport.TabblesServ.serverData GetYellowBlueSoftServer();
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/ITabblesService/GetYellowBlueSoftServer", ReplyAction="http://tempuri.org/ITabblesService/GetYellowBlueSoftServerResponse")]
        System.Threading.Tasks.Task<CrashReport.TabblesServ.serverData> GetYellowBlueSoftServerAsync();
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/ITabblesService/GetNewVersionData", ReplyAction="http://tempuri.org/ITabblesService/GetNewVersionDataResponse")]
        CrashReport.TabblesServ.NewVersionData GetNewVersionData();
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/ITabblesService/GetNewVersionData", ReplyAction="http://tempuri.org/ITabblesService/GetNewVersionDataResponse")]
        System.Threading.Tasks.Task<CrashReport.TabblesServ.NewVersionData> GetNewVersionDataAsync();
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/ITabblesService/SendCrashEmail", ReplyAction="http://tempuri.org/ITabblesService/SendCrashEmailResponse")]
        string SendCrashEmail(string emailBody);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/ITabblesService/SendCrashEmail", ReplyAction="http://tempuri.org/ITabblesService/SendCrashEmailResponse")]
        System.Threading.Tasks.Task<string> SendCrashEmailAsync(string emailBody);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/ITabblesService/SendEmail", ReplyAction="http://tempuri.org/ITabblesService/SendEmailResponse")]
        string SendEmail(string emailBody, string subject);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/ITabblesService/SendEmail", ReplyAction="http://tempuri.org/ITabblesService/SendEmailResponse")]
        System.Threading.Tasks.Task<string> SendEmailAsync(string emailBody, string subject);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/ITabblesService/writeReportToDb", ReplyAction="http://tempuri.org/ITabblesService/writeReportToDbResponse")]
        string writeReportToDb(string stackTrace, string email, System.Nullable<bool> isCloud, string license, string codeVersion, string strCrashId, string windowsUser, string machineName, System.Nullable<bool> crashDialogWasShown);
        
        [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/ITabblesService/writeReportToDb", ReplyAction="http://tempuri.org/ITabblesService/writeReportToDbResponse")]
        System.Threading.Tasks.Task<string> writeReportToDbAsync(string stackTrace, string email, System.Nullable<bool> isCloud, string license, string codeVersion, string strCrashId, string windowsUser, string machineName, System.Nullable<bool> crashDialogWasShown);
    }
    
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
    public interface ITabblesServiceChannel : CrashReport.TabblesServ.ITabblesService, System.ServiceModel.IClientChannel {
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
    public partial class TabblesServiceClient : System.ServiceModel.ClientBase<CrashReport.TabblesServ.ITabblesService>, CrashReport.TabblesServ.ITabblesService {
        
        public TabblesServiceClient() {
        }
        
        public TabblesServiceClient(string endpointConfigurationName) : 
                base(endpointConfigurationName) {
        }
        
        public TabblesServiceClient(string endpointConfigurationName, string remoteAddress) : 
                base(endpointConfigurationName, remoteAddress) {
        }
        
        public TabblesServiceClient(string endpointConfigurationName, System.ServiceModel.EndpointAddress remoteAddress) : 
                base(endpointConfigurationName, remoteAddress) {
        }
        
        public TabblesServiceClient(System.ServiceModel.Channels.Binding binding, System.ServiceModel.EndpointAddress remoteAddress) : 
                base(binding, remoteAddress) {
        }
        
        public CrashReport.TabblesServ.Tabbles4NewVersionData GetTabbles4NewVersionData() {
            return base.Channel.GetTabbles4NewVersionData();
        }
        
        public System.Threading.Tasks.Task<CrashReport.TabblesServ.Tabbles4NewVersionData> GetTabbles4NewVersionDataAsync() {
            return base.Channel.GetTabbles4NewVersionDataAsync();
        }
        
        public CrashReport.TabblesServ.Tabbles4CloudSqlServerData GetTabbles4CloudSqlServerData() {
            return base.Channel.GetTabbles4CloudSqlServerData();
        }
        
        public System.Threading.Tasks.Task<CrashReport.TabblesServ.Tabbles4CloudSqlServerData> GetTabbles4CloudSqlServerDataAsync() {
            return base.Channel.GetTabbles4CloudSqlServerDataAsync();
        }
        
        public CrashReport.TabblesServ.serverData GetYellowBlueSoftServer() {
            return base.Channel.GetYellowBlueSoftServer();
        }
        
        public System.Threading.Tasks.Task<CrashReport.TabblesServ.serverData> GetYellowBlueSoftServerAsync() {
            return base.Channel.GetYellowBlueSoftServerAsync();
        }
        
        public CrashReport.TabblesServ.NewVersionData GetNewVersionData() {
            return base.Channel.GetNewVersionData();
        }
        
        public System.Threading.Tasks.Task<CrashReport.TabblesServ.NewVersionData> GetNewVersionDataAsync() {
            return base.Channel.GetNewVersionDataAsync();
        }
        
        public string SendCrashEmail(string emailBody) {
            return base.Channel.SendCrashEmail(emailBody);
        }
        
        public System.Threading.Tasks.Task<string> SendCrashEmailAsync(string emailBody) {
            return base.Channel.SendCrashEmailAsync(emailBody);
        }
        
        public string SendEmail(string emailBody, string subject) {
            return base.Channel.SendEmail(emailBody, subject);
        }
        
        public System.Threading.Tasks.Task<string> SendEmailAsync(string emailBody, string subject) {
            return base.Channel.SendEmailAsync(emailBody, subject);
        }
        
        public string writeReportToDb(string stackTrace, string email, System.Nullable<bool> isCloud, string license, string codeVersion, string strCrashId, string windowsUser, string machineName, System.Nullable<bool> crashDialogWasShown) {
            return base.Channel.writeReportToDb(stackTrace, email, isCloud, license, codeVersion, strCrashId, windowsUser, machineName, crashDialogWasShown);
        }
        
        public System.Threading.Tasks.Task<string> writeReportToDbAsync(string stackTrace, string email, System.Nullable<bool> isCloud, string license, string codeVersion, string strCrashId, string windowsUser, string machineName, System.Nullable<bool> crashDialogWasShown) {
            return base.Channel.writeReportToDbAsync(stackTrace, email, isCloud, license, codeVersion, strCrashId, windowsUser, machineName, crashDialogWasShown);
        }
    }
}
