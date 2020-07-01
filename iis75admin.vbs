option explicit
'******************************************************************************
'* IIS 7.5 Management Script
'* Version 1.0
'*
'* This script leverages the included class 'clsIIS' to implement a command
'* line utility for retrieving and setting all configurable properties for
'* every object type on IIS 7.5.
'*
'* Usage:	iis7admin get|set propertypath [value]
'*
'* Change Log
'*	Verson 1.0 - 2011.11.27
'*	- first executable; implements the initial requirement
'******************************************************************************

dim objIIS
dim intArgCount
dim strVerb
dim strPath
dim varValue

set objIIS = new clsIIS

intArgCount = wscript.arguments.count

if intArgCount < 2 then
	wscript.echo "ERROR:  Invalid argument count."
	wscript.quit
end if

strVerb = wscript.arguments.item(0)
strPath = wscript.arguments.item(1)

select case strVerb

	case "get"
		if intArgCount <> 2 then
			wscript.echo "ERROR:  Action ""get"" requires only one more argument; property path."
			wscript.quit
		else
			wscript.echo eval("objIIS." & replace(strPath,"'",""""))
		end if

	case "set"
		if intArgCount <> 3 then
			wscript.echo "ERROR:  Action ""set"" requires exactly two more arguments; property path and new value."
			wscript.quit
		else
			varValue = wscript.arguments.item(2)
			execute("objIIS." & replace(strPath,"'","""") & " = " & varValue)
			objIIS.save
		end if

	case else
		wscript.echo "ERROR:  First argument must be either ""get"" or ""set""."
		wscript.quit

end select



'******************************************************************************
'* IIS 7.5 Management Class
'* Version 1.1
'* by Daniel Lee (daniel.lee@drlsystems.com)
'*
'* This set of classes and utility functions implement a simplified interface
'* for the management of IIS 7.5.  Users should only instantiate class 'clsIIS'.
'* The class is self-documenting.  Method 'listall' of 'clsIIS' will echo the
'* complete API to the console.
'*
'*	Usage:		set objIIS = new clsIIS
'*
'* Example:		Echo 'id' property of site "Default Web Site" to the console.
'*						wscript.echo objIIS.site("Default Web Site").id
'*
'* Example:		Set app pool 'DefaultAppPool' property 'autoStart' to 'false'.
'*						wscript.echo objIIS.applicationPool("DefaultAppPool").autoStart = false
'*
'* Change Log
'*	Version 1.0 - 2011.11.27
'*	- first iteration; implements the initial requirement
'*	Version 1.1 - 2011.11.28
'*	- added default get property, boolean exists, to each class, that returns true
'******************************************************************************

'***********
'* IIS
'***********
class clsIIS

	public default property get exists : exists = true : end property
	public property get site ( prm_strSiteName )
		dim pvt_colSites
		dim pvt_objSite
		dim pvt_intSiteIndex

		set pvt_colSites = pvt_objServer.getAdminSection("system.applicationHost/sites", "MACHINE/WEBROOT/APPHOST").collection

		if isObject ( pvt_colSites ) then
			pvt_intSiteIndex = intFindCollectionMemberByPropertyValue ( pvt_colSites, "name", prm_strSiteName )
			if pvt_intSiteIndex > -1 then
				set pvt_objSite = new clsSite
				set pvt_objSite.pub_objSite = pvt_colSites.Item(pvt_intSiteIndex)
			else
				err.raise 8 , "clsIIS" , "Unable to locate site '" & prm_strSiteName & "' on this server.  Site collection at 'system.applicationHost/sites' contains no site with that name."
			end if
		else
			err.raise 8 , "clsIIS" , "IIS server is not currently hosting any web sites.  Site collection at 'system.applicationHost/sites' is empty."
		end if

		set pvt_colSites = nothing
		set site = pvt_objSite
	end property

	public property get applicationPool ( prm_strApplicationPoolName )
		dim pvt_colApplicationPools
		dim pvt_objApplicationPool
		dim pvt_intApplicationPoolIndex

		set pvt_colApplicationPools = pvt_objServer.getAdminSection("system.applicationHost/applicationPools", "MACHINE/WEBROOT/APPHOST").collection

		if isObject ( pvt_colApplicationPools ) then
			pvt_intApplicationPoolIndex = intFindCollectionMemberByPropertyValue ( pvt_colApplicationPools, "name", prm_strApplicationPoolName )
			if pvt_intApplicationPoolIndex > -1 then
				set pvt_objApplicationPool = new clsApplicationPool
				set pvt_objApplicationPool.pub_objApplicationPool = pvt_colApplicationPools.Item(pvt_intApplicationPoolIndex)
			else
				err.raise 8 , "clsIIS" , "Unable to locate application pool '" & prm_strSiteName & "' on this server.  Application pool collection at 'system.applicationHost/applicationPools' contains no application pool with that name."
				set pvt_objApplicationPool = null
			end if
		else
			err.raise 8 , "clsIIS" , "IIS server is not currently hosting any application pools.  Application pool collection at 'system.applicationHost/applicatonPools' is empty."
		end if

		set pvt_colApplicationPools = nothing
		set applicationPool = pvt_objApplicationPool
	end property

	public property get siteDefaults
		dim pvt_colSites
		dim pvt_objSiteDefaults

		set pvt_colSites = pvt_objServer.getAdminSection("system.applicationHost/sites", "MACHINE/WEBROOT/APPHOST")

		if isObject ( pvt_colSites ) then
			set pvt_objSiteDefaults = new clsSiteDefaults
			set pvt_objSiteDefaults.pub_objSiteDefaults = pvt_colSites.childElements.item("siteDefaults")
		else
			err.raise 8 , "clsIIS" , "IIS server is not currently hosting any web sites and therefore has nothing to bind site defaults to.  Site collection at 'system.applicationHost/sites' is empty."
		end if

		set pvt_colSites = nothing
		set siteDefaults = pvt_objSiteDefaults
	end property

	public property get applicationDefaults
		dim pvt_colSites
		dim pvt_objApplicationDefaults

		set pvt_colSites = pvt_objServer.getAdminSection("system.applicationHost/sites", "MACHINE/WEBROOT/APPHOST")

		if isObject ( pvt_colSites ) then
			set pvt_objApplicationDefaults = new clsApplicationDefaults
			set pvt_objApplicationDefaults.pub_objApplicationDefaults = pvt_colSites.childElements.item("applicationDefaults")
		else
			err.raise 8 , "clsIIS" , "IIS server is not currently hosting any web sites and therefore has nothing to application defaults to.  Site collection at 'system.applicationHost/sites' is empty."
		end if

		set pvt_colSites = nothing
		set applicationDefaults = pvt_objApplicationDefaults
	end property

	public property get virtualDirectoryDefaults
		dim pvt_colSites
		dim pvt_objVirtualDirectoryDefaults

		set pvt_colSites = pvt_objServer.getAdminSection("system.applicationHost/sites", "MACHINE/WEBROOT/APPHOST")

		if isObject ( pvt_colSites ) then
			set pvt_objVirtualDirectoryDefaults = new clsVirtualDirectoryDefaults
			set pvt_objVirtualDirectoryDefaults.pub_objVirtualDirectoryDefaults = pvt_colSites.childElements.item("virtualDirectoryDefaults")
		else
			err.raise 8 , "clsIIS" , "IIS server is not currently hosting any web sites and therefore has nothing to bind virtual directory defaults to.  Site collection at 'system.applicationHost/sites' is empty."
		end if

		set pvt_colSites = nothing
		set virtualDirectoryDefaults = pvt_objVirtualDirectoryDefaults
	end property

	public property get applicationPoolDefaults
		dim pvt_colApplicationPools
		dim pvt_objApplicationPoolDefaults

		set pvt_colApplicationPools = pvt_objServer.getAdminSection("system.applicationHost/applicationPools", "MACHINE/WEBROOT/APPHOST")

		if isObject ( pvt_colApplicationPools ) then
			set pvt_objApplicationPoolDefaults = new clsApplicationPoolDefaults
			set pvt_objApplicationPoolDefaults.pub_objApplicationPoolDefaults = pvt_colApplicationPools.childElements.item("applicationPoolDefaults")
		else
			err.raise 8 , "clsIIS" , "IIS server is not currently hosting any application pools and therefore has nothing to bind application pool defaults to.  Application pool collection at 'system.applicationHost/sites' is empty."
		end if

		set pvt_colApplicationPools = nothing
		set applicationPoolDefaults = pvt_objApplicationPoolDefaults
	end property

	private pvt_objServer

   public sub save
		pvt_objServer.commitChanges()
   end sub

   public sub listAll
	   wscript.echo "objIIS.Site(strSiteName).Application(strRelativePath).ApplicationPool"
	   wscript.echo "objIIS.Site(strSiteName).Application(strRelativePath).EnabledProtocols"
	   wscript.echo "objIIS.Site(strSiteName).Application(strRelativePath).Path"
	   wscript.echo "objIIS.Site(strSiteName).Application(strRelativePath).ServiceAutoStartEnabled"
	   wscript.echo "objIIS.Site(strSiteName).Application(strRelativePath).ServiceAutoStartProvider"
	   wscript.echo "objIIS.Site(strSiteName).Application(strRelativePath).VirtualDirectoryDefaults.AllowSubDirConfig"
	   wscript.echo "objIIS.Site(strSiteName).Application(strRelativePath).VirtualDirectoryDefaults.LogonMethod"
	   wscript.echo "objIIS.Site(strSiteName).Application(strRelativePath).VirtualDirectoryDefaults.Password"
	   wscript.echo "objIIS.Site(strSiteName).Application(strRelativePath).VirtualDirectoryDefaults.Path"
	   wscript.echo "objIIS.Site(strSiteName).Application(strRelativePath).VirtualDirectoryDefaults.PhysicalPath"
	   wscript.echo "objIIS.Site(strSiteName).Application(strRelativePath).VirtualDirectoryDefaults.UserName"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).AutoStart"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).CLRConfigFile"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Cpu.Action"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Cpu.Limit"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Cpu.ResetInterval"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Cpu.SmpAffinitized"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Cpu.SmpProcessorAffinityMask"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Cpu.SmpProcessorAffinityMask2"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Enable32BitAppOnWin64"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).EnableConfigurationOverride"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Failure.AutoShutdownExe"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Failure.AutoShutdownParams"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Failure.LoadBalancerCapabilities"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Failure.OrphanActionExe"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Failure.OrphanActionParams"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Failure.OrphanWorkerProcess"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Failure.RapidFailProtection"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Failure.RapidFailProtectionInterval"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Failure.RapidFailProtectionMaxCrashes"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).ManagedPipelineMode"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).ManagedRuntimeLoader"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).ManagedRuntimeVersion"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Name"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).PassAnonymousToken"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).ProcessModel.IdentityType"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).ProcessModel.IdleTimeout"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).ProcessModel.LoadUserProfile"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).ProcessModel.LogonType"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).ProcessModel.ManualGroupMembership"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).ProcessModel.MaxProcesses"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).ProcessModel.Password"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).ProcessModel.PingingEnabled"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).ProcessModel.PingInterval"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).ProcessModel.PingResponseTime"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).ProcessModel.ShutdownTimeLimit"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).ProcessModel.StartupTimeLimit"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).ProcessModel.UserName"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).QueueLength"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Recycling.DisallowOverlappingRotation"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Recycling.DisallowRotationOnConfigChange"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Recycling.LogEventOnRecycle"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Recycling.PeriodicRestart.Memory"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Recycling.PeriodicRestart.PrivateMemory"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Recycling.PeriodicRestart.Requests"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).Recycling.PeriodicRestart.Time"
	   wscript.echo "objIIS.ApplicationPool(strApplicationPoolName).StartMode"
	   wscript.echo "objIIS.ApplicationDefaults.ApplicationPool"
	   wscript.echo "objIIS.ApplicationDefaults.EnabledProtocols"
	   wscript.echo "objIIS.ApplicationDefaults.Path"
	   wscript.echo "objIIS.ApplicationPoolDefaults.AutoStart"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Cpu.Action"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Cpu.Limit"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Cpu.ResetInterval"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Cpu.SmpAffinitized"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Cpu.SmpProcessorAffinityMask"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Cpu.SmpProcessorAffinityMask2"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Enable32BitAppOnWin64"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Failure.AutoShutdownExe"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Failure.AutoShutdownParams"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Failure.LoadBalancerCapabilities"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Failure.OrphanActionExe"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Failure.OrphanActionParams"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Failure.OrphanWorkerProcess"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Failure.RapidFailProtection"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Failure.RapidFailProtectionInterval"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Failure.RapidFailProtectionMaxCrashes"
	   wscript.echo "objIIS.ApplicationPoolDefaults.ManagedPipelineMode"
	   wscript.echo "objIIS.ApplicationPoolDefaults.ManagedRuntimeVersion"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Name"
	   wscript.echo "objIIS.ApplicationPoolDefaults.ProcessModel.IdentityType"
	   wscript.echo "objIIS.ApplicationPoolDefaults.ProcessModel.IdleTimeout"
	   wscript.echo "objIIS.ApplicationPoolDefaults.ProcessModel.LoadUserProfile"
	   wscript.echo "objIIS.ApplicationPoolDefaults.ProcessModel.LogonType"
	   wscript.echo "objIIS.ApplicationPoolDefaults.ProcessModel.ManualGroupMembership"
	   wscript.echo "objIIS.ApplicationPoolDefaults.ProcessModel.MaxProcesses"
	   wscript.echo "objIIS.ApplicationPoolDefaults.ProcessModel.Password"
	   wscript.echo "objIIS.ApplicationPoolDefaults.ProcessModel.PingingEnabled"
	   wscript.echo "objIIS.ApplicationPoolDefaults.ProcessModel.PingInterval"
	   wscript.echo "objIIS.ApplicationPoolDefaults.ProcessModel.PingResponseTime"
	   wscript.echo "objIIS.ApplicationPoolDefaults.ProcessModel.ShutdownTimeLimit"
	   wscript.echo "objIIS.ApplicationPoolDefaults.ProcessModel.StartupTimeLimit"
	   wscript.echo "objIIS.ApplicationPoolDefaults.ProcessModel.UserName"
	   wscript.echo "objIIS.ApplicationPoolDefaults.QueueLength"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Recycling.DisallowOverlappingRotation"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Recycling.DisallowRotationOnConfigChange"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Recycling.LogEventOnRecycle"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Recycling.PeriodicRestart.Memory"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Recycling.PeriodicRestart.PrivateMemory"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Recycling.PeriodicRestart.Requests"
	   wscript.echo "objIIS.ApplicationPoolDefaults.Recycling.PeriodicRestart.Time"
	   wscript.echo "objIIS.ApplicationPoolDefaults.StartMode"
	   wscript.echo "objIIS.SiteDefaults.Binding(strProtocol).BindingInformation"
	   wscript.echo "objIIS.SiteDefaults.Binding(strProtocol).Protocol"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.AllowUTF8"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Connections.ControlChannelTimeout"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Connections.DataChannelTimeout"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Connections.DisableSocketPooling"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Connections.MaxBandwidth"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Connections.MaxConnections"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Connections.MinBytesPerSecond"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Connections.ResetOnMaxConnections"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Connections.ServerListenBacklog"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Connections.UnauthenticatedTimeout"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.DirectoryBrowse.ShowFlags"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.DirectoryBrowse.VirtualDirectoryTimeout"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.FileHandling.AllowReadUploadsInProgress"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.FileHandling.AllowReplaceOnRename"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.FirewallSupport.ExternalIp4Address"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.LogFile.Directory"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.LogFile.Enabled"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.LogFile.LocalTimeRollover"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.LogFile.LogExtFileFlags"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.LogFile.Period"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.LogFile.SelectiveLogging"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.LogFile.TruncateSize"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Messages.AllowLocalDetailedErrors"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Messages.BannerMessage"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Messages.ExitMessage"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Messages.ExpandVariables"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Messages.GreetingMessage"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Messages.MaxClientsMessage"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Messages.SuppressDefaultBanner"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.Authentication.AnonymousAuthentication.DefaultLogonDomain"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.Authentication.AnonymousAuthentication.Enabled"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.Authentication.AnonymousAuthentication.LogonMethod"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.Authentication.AnonymousAuthentication.Password"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.Authentication.AnonymousAuthentication.UserName"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.Authentication.BasicAuthentication.DefaultLogonDomain"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.Authentication.BasicAuthentication.Enabled"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.Authentication.BasicAuthentication.LogonMethod"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.Authentication.ClientCertAuthentication.Enabled"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.CommandFiltering.AllowUnlisted"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.CommandFiltering.MaxCommandLine"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.DataChannelSecurity.MatchClientAddressForPasv"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.DataChannelSecurity.MatchClientAddressForPort"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.Ssl.ControlChannelPolicy"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.Ssl.DataChannelPolicy"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.Ssl.ServerCertHash"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.Ssl.ServerCertStoreName"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.Ssl.Ssl128"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.SslClientCertificates.ClientCertificatePolicy"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.SslClientCertificates.RevocationFreshnessTime"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.SslClientCertificates.RevocationURLRetrievalTimeout"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.SslClientCertificates.UseActiveDirectoryMapping"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.Security.SslClientCertificates.ValidationFlags"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.ServerAutoStart"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.UserIsolation.ActiveDirectory.AdCacheRefresh"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.UserIsolation.ActiveDirectory.AdPassword"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.UserIsolation.ActiveDirectory.AdUserName"
	   wscript.echo "objIIS.SiteDefaults.FtpServer.UserIsolation.Mode"
	   wscript.echo "objIIS.SiteDefaults.Id"
	   wscript.echo "objIIS.SiteDefaults.Limits.ConnectionTimeout"
	   wscript.echo "objIIS.SiteDefaults.Limits.MaxBandwidth"
	   wscript.echo "objIIS.SiteDefaults.Limits.MaxConnections"
	   wscript.echo "objIIS.SiteDefaults.LogFile.CustomLogPluginClsid"
	   wscript.echo "objIIS.SiteDefaults.LogFile.Directory"
	   wscript.echo "objIIS.SiteDefaults.LogFile.Enabled"
	   wscript.echo "objIIS.SiteDefaults.LogFile.LocalTimeRollover"
	   wscript.echo "objIIS.SiteDefaults.LogFile.LogExtFileFlags"
	   wscript.echo "objIIS.SiteDefaults.LogFile.LogFormat"
	   wscript.echo "objIIS.SiteDefaults.LogFile.Period"
	   wscript.echo "objIIS.SiteDefaults.LogFile.TruncateSize"
	   wscript.echo "objIIS.SiteDefaults.Name"
	   wscript.echo "objIIS.SiteDefaults.ServerAutoStart"
	   wscript.echo "objIIS.SiteDefaults.TraceFailedRequestsLogging.CustomActionsEnabled"
	   wscript.echo "objIIS.SiteDefaults.TraceFailedRequestsLogging.Directory"
	   wscript.echo "objIIS.SiteDefaults.TraceFailedRequestsLogging.Enabled"
	   wscript.echo "objIIS.SiteDefaults.TraceFailedRequestsLogging.MaxLogFiles"
	   wscript.echo "objIIS.SiteDefaults.TraceFailedRequestsLogging.MaxLogFileSizeKB"
	   wscript.echo "objIIS.VirtualDirectoryDefaults.AllowSubDirConfig"
	   wscript.echo "objIIS.VirtualDirectoryDefaults.LogonMethod"
	   wscript.echo "objIIS.VirtualDirectoryDefaults.Password"
	   wscript.echo "objIIS.VirtualDirectoryDefaults.Path"
	   wscript.echo "objIIS.VirtualDirectoryDefaults.PhysicalPath"
	   wscript.echo "objIIS.VirtualDirectoryDefaults.UserName"
	   wscript.echo "objIIS.Site(strSiteName).ApplicationDefaults.ApplicationPool"
	   wscript.echo "objIIS.Site(strSiteName).ApplicationDefaults.EnabledProtocols"
	   wscript.echo "objIIS.Site(strSiteName).ApplicationDefaults.Path"
	   wscript.echo "objIIS.Site(strSiteName).ApplicationDefaults.ServiceAutoStartEnabled"
	   wscript.echo "objIIS.Site(strSiteName).ApplicationDefaults.ServiceAutoStartProvider"
	   wscript.echo "objIIS.Site(strSiteName).Binding(strProtocol).BindingInformation"
	   wscript.echo "objIIS.Site(strSiteName).Binding(strProtocol).Protocol"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.AllowUTF8"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Connections.ControlChannelTimeout"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Connections.DataChannelTimeout"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Connections.DisableSocketPooling"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Connections.MaxBandwidth"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Connections.MaxConnections"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Connections.MinBytesPerSecond"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Connections.ResetOnMaxConnections"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Connections.ServerListenBacklog"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Connections.UnauthenticatedTimeout"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.DirectoryBrowse.ShowFlags"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.DirectoryBrowse.VirtualDirectoryTimeout"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.FileHandling.AllowReadUploadsInProgress"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.FileHandling.AllowReplaceOnRename"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.FirewallSupport.ExternalIp4Address"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.LogFile.Directory"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.LogFile.Enabled"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.LogFile.LocalTimeRollover"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.LogFile.LogExtFileFlags"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.LogFile.Period"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.LogFile.SelectiveLogging"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.LogFile.TruncateSize"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Messages.AllowLocalDetailedErrors"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Messages.BannerMessage"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Messages.ExitMessage"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Messages.ExpandVariables"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Messages.GreetingMessage"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Messages.MaxClientsMessage"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Messages.SuppressDefaultBanner"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.Authentication.AnonymousAuthentication.DefaultLogonDomain"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.Authentication.AnonymousAuthentication.Enabled"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.Authentication.AnonymousAuthentication.LogonMethod"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.Authentication.AnonymousAuthentication.Password"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.Authentication.AnonymousAuthentication.UserName"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.Authentication.BasicAuthentication.DefaultLogonDomain"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.Authentication.BasicAuthentication.Enabled"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.Authentication.BasicAuthentication.LogonMethod"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.Authentication.ClientCertAuthentication.Enabled"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.CommandFiltering.AllowUnlisted"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.CommandFiltering.MaxCommandLine"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.DataChannelSecurity.MatchClientAddressForPasv"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.DataChannelSecurity.MatchClientAddressForPort"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.Ssl.ControlChannelPolicy"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.Ssl.DataChannelPolicy"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.Ssl.ServerCertHash"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.Ssl.ServerCertStoreName"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.Ssl.Ssl128"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.SslClientCertificates.ClientCertificatePolicy"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.SslClientCertificates.RevocationFreshnessTime"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.SslClientCertificates.RevocationURLRetrievalTimeout"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.SslClientCertificates.UseActiveDirectoryMapping"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.Security.SslClientCertificates.ValidationFlags"
	   wscript.echo "objIIS.Site(strSiteName).FtpServer.ServerAutoStart"
	   wscript.echo "objIIS.Site(strSiteName).Id"
	   wscript.echo "objIIS.Site(strSiteName).Limits.ConnectionTimeout"
	   wscript.echo "objIIS.Site(strSiteName).Limits.MaxBandwidth"
	   wscript.echo "objIIS.Site(strSiteName).Limits.MaxConnections"
	   wscript.echo "objIIS.Site(strSiteName).LogFile.CustomLogPluginClsid"
	   wscript.echo "objIIS.Site(strSiteName).LogFile.Directory"
	   wscript.echo "objIIS.Site(strSiteName).LogFile.Enabled"
	   wscript.echo "objIIS.Site(strSiteName).LogFile.LocalTimeRollover"
	   wscript.echo "objIIS.Site(strSiteName).LogFile.LogExtFileFlags"
	   wscript.echo "objIIS.Site(strSiteName).LogFile.LogFormat"
	   wscript.echo "objIIS.Site(strSiteName).LogFile.Period"
	   wscript.echo "objIIS.Site(strSiteName).LogFile.TruncateSize"
	   wscript.echo "objIIS.Site(strSiteName).Name"
	   wscript.echo "objIIS.Site(strSiteName).ServerAutoStart"
	   wscript.echo "objIIS.Site(strSiteName).TraceFailedRequestsLogging.CustomActionsEnabled"
	   wscript.echo "objIIS.Site(strSiteName).TraceFailedRequestsLogging.Directory"
	   wscript.echo "objIIS.Site(strSiteName).TraceFailedRequestsLogging.Enabled"
	   wscript.echo "objIIS.Site(strSiteName).TraceFailedRequestsLogging.MaxLogFiles"
	   wscript.echo "objIIS.Site(strSiteName).TraceFailedRequestsLogging.MaxLogFileSizeKB"
	   wscript.echo "objIIS.Site(strSiteName).VirtualDirectoryDefaults.AllowSubDirConfig"
	   wscript.echo "objIIS.Site(strSiteName).VirtualDirectoryDefaults.LogonMethod"
	   wscript.echo "objIIS.Site(strSiteName).VirtualDirectoryDefaults.Password"
	   wscript.echo "objIIS.Site(strSiteName).VirtualDirectoryDefaults.Path"
	   wscript.echo "objIIS.Site(strSiteName).VirtualDirectoryDefaults.PhysicalPath"
	   wscript.echo "objIIS.Site(strSiteName).VirtualDirectoryDefaults.UserName"
	   wscript.echo "objIIS.Site(strSiteName).Application(strRelativePath).virtualDirectory(strRelativePath).AllowSubDirConfig"
	   wscript.echo "objIIS.Site(strSiteName).Application(strRelativePath).virtualDirectory(strRelativePath).LogonMethod"
	   wscript.echo "objIIS.Site(strSiteName).Application(strRelativePath).virtualDirectory(strRelativePath).Password"
	   wscript.echo "objIIS.Site(strSiteName).Application(strRelativePath).virtualDirectory(strRelativePath).Path"
	   wscript.echo "objIIS.Site(strSiteName).Application(strRelativePath).virtualDirectory(strRelativePath).PhysicalPath"
	   wscript.echo "objIIS.Site(strSiteName).Application(strRelativePath).virtualDirectory(strRelativePath).UserName"
   end sub

	sub class_Initialize
		on error resume next
			set pvt_objServer = wscript.createObject("Microsoft.ApplicationHost.WritableAdminManager")
		on error goto 0
		if isObject(pvt_objServer) then
			pvt_objServer.CommitPath = "MACHINE/WEBROOT/APPHOST"
		else
			err.raise 8 , "clsIIS" , "Unable to connect to IIS on this server.  Could not create an instance of 'Microsoft.ApplicationHost.WritableAdminManager'."
		end if
	end sub

	sub class_Terminate
		set pvt_objServer = nothing
	end sub

end class

'************
'* Site
'************
class clsSite

	public default property get exists : exists = true : end property
	public property get id : id = pub_objSite.properties.item("id").value : end property
	public property get name : name = pub_objSite.properties.item("name").value : end property
	public property get serverAutoStart : serverAutoStart = pub_objSite.properties.item("serverAutoStart").value : end property

	public property let id ( prm_id ) : pub_objSite.properties.item("id").value = prm_id : end property
	public property let name ( prm_name ) : pub_objSite.properties.item("name").value = prm_name : end property
	public property let serverAutoStart ( prm_serverAutoStart ) : pub_objSite.properties.item("serverAutoStart").value = prm_serverAutoStart : end property

   public property get binding ( prm_strProtocol )
		dim pvt_colBindings
		dim pvt_objBinding
		dim pvt_intBindingIndex

      set pvt_colBindings = pub_objSite.childElements.item("bindings").collection

  		if isObject ( pvt_colBindings ) then
			pvt_intBindingIndex = intFindCollectionMemberByPropertyValue ( pvt_colBindings, "protocol", prm_strProtocol )
			if pvt_intBindingIndex > -1 then
   	      set pvt_objBinding = new clsBinding
   	      set pvt_objBinding.pub_objBinding = pvt_colBindings.item(pvt_intBindingIndex)
 			else
				err.raise 8 , "clsSite" , "Unable to locate binding information for protocol '" & prm_strProtocol & "' in site '" & me.name & "'.  Bindings collection for this site object contains no binding to that protocol."
			end if
		else
			err.raise 8 , "clsSite" , "Site '" & me.name & "' is not currently bound to any protocol.  Bindings collection for this site is empty."
		end if

      set pvt_colBindings = nothing
		set binding = pvt_objBinding
	end property

   public property get applicationDefaults
		dim pvt_objApplicationDefaults
		set pvt_objApplicationDefaults = new clsApplicationDefaults
		set pvt_objApplicationDefaults.pub_objApplicationDefaults = pub_objSite.childElements.item("applicationDefaults")
		set applicationDefaults = pvt_objApplicationDefaults
	end property

   public property get ftpServer
		dim pvt_objFtpServer
		set pvt_objFtpServer = new clsFtpServer
		set pvt_objFtpServer.pub_objFtpServer = pub_objSite.childElements.item("ftpServer")
		set ftpServer = pvt_objFtpServer
	end property

   public property get limits
		dim pvt_objLimits
		set pvt_objLimits = new clsLimits
		set pvt_objLimits.pub_objLimits = pub_objSite.childElements.item("limits")
		set limits = pvt_objLimits
	end property

   public property get logFile
		dim pvt_objLogFile
		set pvt_objLogFile = new clsLogFile
		set pvt_objLogFile.pub_objLogFile = pub_objSite.childElements.item("logFile")
		set logFile = pvt_objLogFile
	end property

   public property get traceFailedRequestsLogging
		dim pvt_objTraceFailedRequestsLogging
		set pvt_objTraceFailedRequestsLogging = new clsTraceFailedRequestsLogging
		set pvt_objTraceFailedRequestsLogging.pub_objTraceFailedRequestsLogging = pub_objSite.childElements.item("traceFailedRequestsLogging")
		set traceFailedRequestsLogging = pvt_objTraceFailedRequestsLogging
	end property

   public property get virtualDirectoryDefaults
		dim pvt_objVirtualDirectoryDefaults
		set pvt_objVirtualDirectoryDefaults = new clsVirtualDirectoryDefaults
		set pvt_objVirtualDirectoryDefaults.pub_objVirtualDirectoryDefaults = pub_objSite.childElements.item("virtualDirectoryDefaults")
		set virtualDirectoryDefaults = pvt_objVirtualDirectoryDefaults
	end property

   public property get application ( prm_strRelativePath )
		dim pvt_colApplications
		dim pvt_objApplication
		dim pvt_intApplicationIndex

		set pvt_colApplications = pub_objSite.collection

		if isObject ( pvt_colApplications ) then
			pvt_intApplicationIndex = intFindCollectionMemberByPropertyValue ( pvt_colApplications, "path", prm_strRelativePath )
			if pvt_intApplicationIndex > -1 then
				set pvt_objApplication = new clsApplication
				set pvt_objApplication.pub_objApplication = pvt_colApplications.Item(pvt_intApplicationIndex)
			else
				err.raise 8 , "clsSite" , "Unable to locate application at path '" & prm_strRelativePath & "' in site '" & me.name & "'.  Application collection for this site object contains no application at that path."
			end if
		else
			err.raise 8 , "clsSite" , "Site '" & me.name & "' is not currently hosting any applications.  Application collection for this site is empty."
		end if

		set pvt_colApplications = nothing
		set application = pvt_objApplication
	end property

   public pub_objSite

end class


'**********************************
'* Trace Failed Requests Logging
'**********************************
class clsTraceFailedRequestsLogging
	public default property get exists : exists = true : end property
  	public property get customActionsEnabled : customActionsEnabled = pub_objTraceFailedRequestsLogging.properties.item("customActionsEnabled").value : end property
  	public property get directory : directory = pub_objTraceFailedRequestsLogging.properties.item("directory").value : end property
  	public property get enabled : enabled = pub_objTraceFailedRequestsLogging.properties.item("enabled").value : end property
  	public property get maxLogFiles : maxLogFiles = pub_objTraceFailedRequestsLogging.properties.item("maxLogFiles").value : end property
  	public property get maxLogFileSizeKB : maxLogFileSizeKB = pub_objTraceFailedRequestsLogging.properties.item("maxLogFileSizeKB").value : end property

  	public property let customActionsEnabled ( prm_customActionsEnabled ) : pub_objTraceFailedRequestsLogging.properties.item("customActionsEnabled").value = prm_customActionsEnabled : end property
  	public property let directory ( prm_directory ) : pub_objTraceFailedRequestsLogging.properties.item("directory").value = prm_directory : end property
  	public property let enabled ( prm_enabled ) : pub_objTraceFailedRequestsLogging.properties.item("enabled").value = prm_enabled : end property
  	public property let maxLogFiles ( prm_maxLogFiles ) : pub_objTraceFailedRequestsLogging.properties.item("maxLogFiles").value = prm_maxLogFiles : end property
  	public property let maxLogFileSizeKB ( prm_maxLogFileSizeKB ) : pub_objTraceFailedRequestsLogging.properties.item("maxLogFileSizeKB").value = prm_maxLogFileSizeKB : end property

	public pub_objTraceFailedRequestsLogging
end class

'**************
'* Limits
'**************
class clsLimits
	public default property get exists : exists = true : end property
	public property get connectionTimeout : connectionTimeout = pub_objLimits.properties.item("connectionTimeout").value : end property
	public property get maxBandwidth : maxBandwidth = pub_objLimits.properties.item("maxBandwidth").value : end property
	public property get maxConnections : maxConnections = pub_objLimits.properties.item("maxConnections").value : end property

   public property let connectionTimeout ( prm_connectionTimeout ) : pub_objLimits.properties.item("connectionTimeout").value = prm_connectionTimeout : end property
	public property let maxBandwidth ( prm_maxBandwidth ) : pub_objLimits.properties.item("maxBandwidth").value = prm_maxBandwidth : end property
	public property let maxConnections ( prm_maxConnections ) : pub_objLimits.properties.item("maxConnections").value = prm_maxConnections : end property

	public pub_objLimits
end class


'***********************
'* Application Pool
'***********************
class clsApplicationPool
	public default property get exists : exists = true : end property
	public property get autoStart : autoStart = pub_objApplicationPool.properties.item("autoStart").value : end property
	public property get clrConfigFile : clrConfigFile = pub_objApplicationPool.properties.item("clrConfigFile").value : end property
	public property get enable32BitAppOnWin64 : enable32BitAppOnWin64 = pub_objApplicationPool.properties.item("enable32BitAppOnWin64").value : end property
	public property get enableConfigurationOverride : enableConfigurationOverride = pub_objApplicationPool.properties.item("enableConfigurationOverride").value : end property
	public property get managedPipelineMode : managedPipelineMode = pub_objApplicationPool.properties.item("managedPipelineMode").value : end property
	public property get managedRuntimeLoader : managedRuntimeLoader = pub_objApplicationPool.properties.item("managedRuntimeLoader").value : end property
	public property get managedRuntimeVersion : managedRuntimeVersion = pub_objApplicationPool.properties.item("managedRuntimeVersion").value : end property
	public property get name : name = pub_objApplicationPool.properties.item("name").value : end property
	public property get passAnonymousToken : passAnonymousToken = pub_objApplicationPool.properties.item("passAnonymousToken").value : end property
	public property get queueLength : queueLength = pub_objApplicationPool.properties.item("queueLength").value : end property
	public property get startMode : startMode = pub_objApplicationPool.properties.item("startMode").value : end property

   public property let autoStart ( prm_autoStart ) : pub_objApplicationPool.properties.item("autoStart").value = prm_autoStart : end property
	public property let clrConfigFile ( prm_clrConfigFile ) : pub_objApplicationPool.properties.item("clrConfigFile").value = prm_clrConfigFile : end property
	public property let enable32BitAppOnWin64 ( prm_enable32BitAppOnWin64 ) : pub_objApplicationPool.properties.item("enable32BitAppOnWin64").value = prm_enable32BitAppOnWin64 : end property
	public property let enableConfigurationOverride ( prm_enableConfigurationOverride ) : pub_objApplicationPool.properties.item("enableConfigurationOverride").value = prm_enableConfigurationOverride : end property
	public property let managedPipelineMode ( prm_managedPipelineMode ) : pub_objApplicationPool.properties.item("managedPipelineMode").value = prm_managedPipelineMode : end property
	public property let managedRuntimeLoader ( prm_managedRuntimeLoader ) : pub_objApplicationPool.properties.item("managedRuntimeLoader").value = prm_managedRuntimeLoader : end property
	public property let managedRuntimeVersion ( prm_managedRuntimeVersion ) : pub_objApplicationPool.properties.item("managedRuntimeVersion").value = prm_managedRuntimeVersion : end property
	public property let name ( prm_name ) : pub_objApplicationPool.properties.item("name").value = prm_name : end property
	public property let passAnonymousToken ( prm_passAnonymousToken ) : pub_objApplicationPool.properties.item("passAnonymousToken").value = prm_passAnonymousToken : end property
	public property let queueLength ( prm_queueLength ) : pub_objApplicationPool.properties.item("queueLength").value = prm_queueLength : end property
	public property let startMode ( prm_startMode ) : pub_objApplicationPool.properties.item("startMode").value = prm_startMode : end property

	public property get cpu
		dim pvt_objCpu
		set pvt_objCpu = new clsCpu
		set pvt_objCpu.pub_objCpu = pub_objApplicationPool.childElements.item("cpu")
		set cpu = pvt_objCpu
	end property

   public property get failure
		dim pvt_objFailure
		set pvt_objFailure = new clsFailure
		set pvt_objFailure.pub_objFailure = pub_objApplicationPool.childElements.item("failure")
		set failure = pvt_objFailure
	end property

   public property get processModel
		dim pvt_objProcessModel
		set pvt_objProcessModel = new clsProcessModel
		set pvt_objProcessModel.pub_objProcessModel = pub_objApplicationPool.childElements.item("processModel")
		set processModel = pvt_objProcessModel
	end property

   public property get recycling
		dim pvt_objRecycling
		set pvt_objRecycling = new clsRecycling
		set pvt_objRecycling.pub_objRecycling = pub_objApplicationPool.childElements.item("recycling")
		set recycling = pvt_objRecycling
	end property

	public pub_objApplicationPool
end class

'********************
'* Site Defaults
'********************
class clsSiteDefaults
	public default property get exists : exists = true : end property
	public property get id : id = pub_objSiteDefaults.properties.item("id").value : end property
	public property get name : name = pub_objSiteDefaults.properties.item("name").value : end property
	public property get serverAutoStart : serverAutoStart = pub_objSiteDefaults.properties.item("serverAutoStart").value : end property

  	public property let id ( prm_id ) : pub_objSiteDefaults.properties.item("id").value = prm_id : end property
	public property let name ( prm_name ) : pub_objSiteDefaults.properties.item("name").value = prm_name : end property
	public property let serverAutoStart ( prm_serverAutoStart ) : pub_objSiteDefaults.properties.item("serverAutoStart").value = prm_serverAutoStart : end property

   public property get binding ( prm_strProtocol )
		dim pvt_colBindings
		dim pvt_objBinding
		dim pvt_intBindingIndex

      set pvt_colBindings = pub_objSiteDefaults.childElements.item("bindings").collection

  		if isObject ( pvt_colBindings ) then
			pvt_intBindingIndex = intFindCollectionMemberByPropertyValue ( pvt_colBindings, "protocol", prm_strProtocol )
			if pvt_intApplicationIndex > -1 then
   	      set pvt_objBinding = new clsBinding
   	      set pvt_objBinding.pub_objBinding = pvt_colBindings.item(pvt_intApplicationIndex)
 			else
				err.raise 8 , "clsSite" , "Unable to locate binding information for protocol '" & prm_strProtocol & "' in this server's site defaults.  Bindings collection for this server's site defaults object contains no binding to that protocol."
			end if
		else
			err.raise 8 , "clsSite" , "Site defaults for this server do not include any protocol bindings.  Bindings collection for this server's site defaults is empty."
		end if

      set pvt_colBindings = nothing
		set bindings = pvt_objBindings
	end property

   public property get ftpServer
		dim pvt_objFtpServer
		set pvt_objFtpServer = new clsFtpServer
		set pvt_objFtpServer.pub_objFtpServer = pub_objSiteDefaults.childElements.item("ftpServer")
		set ftpServer = pvt_objFtpServer
	end property

   public property get limits
		dim pvt_objLimits
		set pvt_objLimits = new clsLimits
		set pvt_objLimits.pub_objLimits = pub_objSiteDefaults.childElements.item("limits")
		set limits = pvt_objLimits
	end property

   public property get logFile
		dim pvt_objLogFile
		set pvt_objLogFile = new clsLogFile
		set pvt_objLogFile.pub_objLogFile = pub_objSiteDefaults.childElements.item("logFile")
		set logFile = pvt_objLogFile
	end property

   public property get traceFailedRequestsLogging
		dim pvt_objTraceFailedRequestsLogging
		set pvt_objTraceFailedRequestsLogging = new clsTraceFailedRequestsLogging
		set pvt_objTraceFailedRequestsLogging.pub_objTraceFailedRequestsLogging = pub_objSiteDefaults.childElements.item("traceFailedRequestsLogging")
		set traceFailedRequestsLogging = pvt_objTraceFailedRequestsLogging
	end property

   public property get virtualDirectoryDefaults
		dim pvt_objVirtualDirectoryDefaults
		set pvt_objVirtualDirectoryDefaults = new clsVirtualDirectoryDefaults
		set pvt_objVirtualDirectoryDefaults.pub_objVirtualDirectoryDefaults = pub_objSiteDefaults.childElements.item("virtualDirectoryDefaults")
		set virtualDirectoryDefaults = pvt_objVirtualDirectoryDefaults
	end property

	public pub_objSiteDefaults
end class

'***************************
'* Application Defaults
'***************************
class clsApplicationDefaults
	public default property get exists : exists = true : end property
	public property get applicationPool : applicationPool = pub_objApplicationDefaults.properties.item("applicationPool").value : end property
	public property get enabledProtocols : enabledProtocols = pub_objApplicationDefaults.properties.item("enabledProtocols").value : end property
	public property get path : path = pub_objApplicationDefaults.properties.item("path").value : end property
	public property get serviceAutoStartEnabled : serviceAutoStartEnabled = pub_objApplicationDefaults.properties.item("serviceAutoStartEnabled").value : end property
	public property get serviceAutoStartProvider : serviceAutoStartProvider = pub_objApplicationDefaults.properties.item("serviceAutoStartProvider").value : end property

  	public property let applicationPool ( prm_applicationPool ) : pub_objApplicationDefaults.properties.item("applicationPool").value = prm_applicationPool : end property
	public property let enabledProtocols ( prm_enabledProtocols ) : pub_objApplicationDefaults.properties.item("enabledProtocols").value = prm_enabledProtocols : end property
	public property let path ( prm_path ) : pub_objApplicationDefaults.properties.item("path").value = prm_path : end property
	public property let serviceAutoStartEnabled ( prm_serviceAutoStartEnabled ) : pub_objApplicationDefaults.properties.item("serviceAutoStartEnabled").value = prm_serviceAutoStartEnabled : end property
	public property let serviceAutoStartProvider ( prm_serviceAutoStartProvider ) : pub_objApplicationDefaults.properties.item("serviceAutoStartProvider").value = prm_serviceAutoStartProvider : end property

	public pub_objApplicationDefaults
end class

'********************************
'* Virtual Directory Defaults
'********************************
class clsVirtualDirectoryDefaults
	public default property get exists : exists = true : end property
	public property get allowSubDirConfig : allowSubDirConfig = pub_objVirtualDirectoryDefaults.properties.item("allowSubDirConfig").value : end property
	public property get logonMethod : logonMethod = pub_objVirtualDirectoryDefaults.properties.item("logonMethod").value : end property
	public property get password : password = pub_objVirtualDirectoryDefaults.properties.item("password").value : end property
	public property get path : path = pub_objVirtualDirectoryDefaults.properties.item("path").value : end property
	public property get physicalPath : physicalPath = pub_objVirtualDirectoryDefaults.properties.item("physicalPath").value : end property
	public property get userName : userName = pub_objVirtualDirectoryDefaults.properties.item("userName").value : end property

	public property let allowSubDirConfig ( prm_allowSubDirConfig ) : pub_objVirtualDirectoryDefaults.properties.item("allowSubDirConfig").value = prm_allowSubDirConfig : end property
	public property let logonMethod ( prm_logonMethod ) : pub_objVirtualDirectoryDefaults.properties.item("logonMethod").value = prm_logonMethod : end property
	public property let password ( prm_password ) : pub_objVirtualDirectoryDefaults.properties.item("password").value = prm_password : end property
	public property let path ( prm_path ) : pub_objVirtualDirectoryDefaults.properties.item("path").value = prm_path : end property
	public property let physicalPath ( prm_physicalPath ) : pub_objVirtualDirectoryDefaults.properties.item("physicalPath").value = prm_physicalPath : end property
	public property let userName ( prm_userName ) : pub_objVirtualDirectoryDefaults.properties.item("userName").value = prm_userName : end property

	public pub_objVirtualDirectoryDefaults
end class

'*******************************
'* Application Pool Defaults
'*******************************
class clsApplicationPoolDefaults
	public default property get exists : exists = true : end property
	public property get autoStart : autoStart = pub_objApplicationPoolDefaults.properties.item("autoStart").value : end property
	public property get clrConfigFile : clrConfigFile = pub_objApplicationPoolDefaults.properties.item("clrConfigFile").value : end property
	public property get enable32BitAppOnWin64 : enable32BitAppOnWin64 = pub_objApplicationPoolDefaults.properties.item("enable32BitAppOnWin64").value : end property
	public property get managedPipelineMode : managedPipelineMode = pub_objApplicationPoolDefaults.properties.item("managedPipelineMode").value : end property
	public property get managedRuntimeLoader : managedRuntimeLoader = pub_objApplicationPoolDefaults.properties.item("managedRuntimeLoader").value : end property
	public property get managedRuntimeVersion : managedRuntimeVersion = pub_objApplicationPoolDefaults.properties.item("managedRuntimeVersion").value : end property
	public property get name : name = pub_objApplicationPoolDefaults.properties.item("name").value : end property
	public property get queueLength : queueLength = pub_objApplicationPoolDefaults.properties.item("queueLength").value : end property
	public property get startMode : startMode = pub_objApplicationPoolDefaults.properties.item("startMode").value : end property

   public property let autoStart ( prm_autoStart ) : pub_objApplicationPoolDefaults.properties.item("autoStart").value = prm_autoStart : end property
	public property let clrConfigFile ( prm_clrConfigFile ) : pub_objApplicationPoolDefaults.properties.item("clrConfigFile").value = prm_clrConfigFile : end property
	public property let enable32BitAppOnWin64 ( prm_enable32BitAppOnWin64 ) : pub_objApplicationPoolDefaults.properties.item("enable32BitAppOnWin64").value = prm_enable32BitAppOnWin64 : end property
	public property let managedPipelineMode ( prm_managedPipelineMode ) : pub_objApplicationPoolDefaults.properties.item("managedPipelineMode").value = prm_managedPipelineMode : end property
	public property let managedRuntimeLoader ( prm_managedRuntimeLoader ) : pub_objApplicationPoolDefaults.properties.item("managedRuntimeLoader").value = prm_managedRuntimeLoader : end property
	public property let managedRuntimeVersion ( prm_managedRuntimeVersion ) : pub_objApplicationPoolDefaults.properties.item("managedRuntimeVersion").value = prm_managedRuntimeVersion : end property
	public property let name ( prm_name ) : pub_objApplicationPoolDefaults.properties.item("name").value = prm_name : end property
	public property let queueLength ( prm_queueLength ) : pub_objApplicationPoolDefaults.properties.item("queueLength").value = prm_queueLength : end property
	public property let startMode ( prm_startMode ) : pub_objApplicationPoolDefaults.properties.item("startMode").value = prm_startMode : end property

	public property get cpu
		dim pvt_objCpu
		set pvt_objCpu = new clsCpu
		set pvt_objCpu.pub_objCpu = pub_objApplicationPoolDefaults.childElements.item("cpu")
		set cpu = pvt_objCpu
	end property

   public property get failure
		dim pvt_objFailure
		set pvt_objFailure = new clsFailure
		set pvt_objFailure.pub_objFailure = pub_objApplicationPoolDefaults.childElements.item("failure")
		set failure = pvt_objFailure
	end property

   public property get processModel
		dim pvt_objProcessModel
		set pvt_objProcessModel = new clsProcessModel
		set pvt_objProcessModel.pub_objProcessModel = pub_objApplicationPoolDefaults.childElements.item("processModel")
		set processModel = pvt_objProcessModel
	end property

   public property get recycling
		dim pvt_objRecycling
		set pvt_objRecycling = new clsRecycling
		set pvt_objRecycling.pub_objRecycling = pub_objApplicationPoolDefaults.childElements.item("recycling")
		set recycling = pvt_objRecycling
	end property

	public pub_objApplicationPoolDefaults
end class

'*******************
'* Application
'*******************
class clsApplication
	public default property get exists : exists = true : end property
	public property get applicationPool : applicationPool = pub_objApplication.properties.item("applicationPool").value : end property
	public property get enabledProtocols : enabledProtocols = pub_objApplication.properties.item("enabledProtocols").value : end property
	public property get path : path = pub_objApplication.properties.item("path").value : end property
	public property get serviceAutoStartEnabled : serviceAutoStartEnabled = pub_objApplication.properties.item("serviceAutoStartEnabled").value : end property
	public property get serviceAutoStartProvider : serviceAutoStartProvider = pub_objApplication.properties.item("serviceAutoStartProvider").value : end property

	public property let applicationPool ( prm_applicationPool ) : pub_objApplication.properties.item("applicationPool").value = prm_applicationPool : end property
	public property let enabledProtocols ( prm_enabledProtocols ) : pub_objApplication.properties.item("enabledProtocols").value = prm_enabledProtocols : end property
	public property let path ( prm_path ) : pub_objApplication.properties.item("path").value = prm_path : end property
	public property let serviceAutoStartEnabled ( prm_serviceAutoStartEnabled ) : pub_objApplication.properties.item("serviceAutoStartEnabled").value = prm_serviceAutoStartEnabled : end property
	public property let serviceAutoStartProvider ( prm_serviceAutoStartProvider ) : pub_objApplication.properties.item("serviceAutoStartProvider").value = prm_serviceAutoStartProvider : end property

	public property get virtualDirectoryDefaults
		dim pvt_objVirtualDirectoryDefaults
		set pvt_objVirtualDirectoryDefaults = new clsVirtualDirectoryDefaults
		set pvt_objVirtualDirectoryDefaults.pub_objVirtualDirectoryDefaults = pub_objApplication.childElements.item("virtualDirectoryDefaults")
		set virtualDirectoryDefaults = pvt_objVirtualDirectoryDefaults
	end property

	public property get virtualDirectory ( prm_strRelativePath )
		dim pvt_colVirtualDirectories
		dim pvt_objVirtualDirectory
		dim pvt_intVirtualDirectoryIndex

		set pvt_colVirtualDirectories = pub_objApplication.collection

		if isObject ( pvt_colVirtualDirectories ) then
			pvt_intVirtualDirectoryIndex = intFindCollectionMemberByPropertyValue ( pvt_colVirtualDirectories, "path", prm_strRelativePath )
			if pvt_intVirtualDirectoryIndex > -1 then
				set pvt_objVirtualDirectory = new clsVirtualDirectory
				set pvt_objVirtualDirectory.pub_objVirtualDirectory = pvt_colVirtualDirectories.Item(pvt_intVirtualDirectoryIndex)
			else
				err.raise 8 , "clsApplication" , "Unable to locate virtual directory '" & prm_strRelativePath & "' in application '" & me.path & "'.  Virtual directories collection for this application object contains no virtual directory at that path."
			end if
		else
			err.raise 8 , "clsApplication" , "Application '" & me.path & "' does not currently contain any virtual directories.  Virtual directory collection for this application is empty."
		end if

		set pvt_colVirtualDirectories = nothing
		set virtualDirectory = pvt_objVirtualDirectory
	end property

   public pub_objApplication
end class

'***************
'* Binding
'***************
class clsBinding
	public default property get exists : exists = true : end property
	public property get bindingInformation : bindingInformation = pub_objBinding.properties.item("bindingInformation").value : end property
	public property get protocol : protocol = pub_objBinding.properties.item("protocol").value : end property

  	public property let bindingInformation ( prm_bindingInformation ) : pub_objBinding.properties.item("bindingInformation").value = prm_bindingInformation : end property
	public property let protocol ( prm_protocol ) : pub_objBinding.properties.item("protocol").value = prm_protocol : end property

	public pub_objBinding
end class

'************************
'* Virtual Directory
'************************
class clsVirtualDirectory
	public default property get exists : exists = true : end property
	public property get allowSubDirConfig : allowSubDirConfig = pub_objVirtualDirectory.properties.item("allowSubDirConfig").value : end property
	public property get logonMethod : logonMethod = pub_objVirtualDirectory.properties.item("logonMethod").value : end property
	public property get password : password = pub_objVirtualDirectory.properties.item("password").value : end property
	public property get path : path = pub_objVirtualDirectory.properties.item("path").value : end property
	public property get physicalPath : physicalPath = pub_objVirtualDirectory.properties.item("physicalPath").value : end property
	public property get userName : userName = pub_objVirtualDirectory.properties.item("userName").value : end property

   public property let allowSubDirConfig ( prm_allowSubDirConfig ) : pub_objVirtualDirectory.properties.item("allowSubDirConfig").value = prm_allowSubDirConfig : end property
	public property let logonMethod ( prm_logonMethod ) : pub_objVirtualDirectory.properties.item("logonMethod").value = prm_logonMethod : end property
	public property let password ( prm_password ) : pub_objVirtualDirectory.properties.item("password").value = prm_password : end property
	public property let path ( prm_path ) : pub_objVirtualDirectory.properties.item("path").value = prm_path : end property
	public property let physicalPath ( prm_physicalPath ) : pub_objVirtualDirectory.properties.item("physicalPath").value = prm_physicalPath : end property
	public property let userName ( prm_userName ) : pub_objVirtualDirectory.properties.item("userName").value = prm_userName : end property

	public pub_objVirtualDirectory
end class

'***********
'* CPU
'***********
class clsCPU
	public default property get exists : exists = true : end property
	public property get action : action = pub_objCPU.properties.item("action").value : end property
	public property get limit : limit = pub_objCPU.properties.item("limit").value : end property
	public property get resetInterval : resetInterval = pub_objCPU.properties.item("resetInterval").value : end property
	public property get smpAffinitized : smpAffinitized = pub_objCPU.properties.item("smpAffinitized").value : end property
	public property get smpProcessorAffinityMask : smpProcessorAffinityMask = pub_objCPU.properties.item("smpProcessorAffinityMask").value : end property
	public property get smpProcessorAffinityMask2 : smpProcessorAffinityMask2 = pub_objCPU.properties.item("smpProcessorAffinityMask2").value : end property

   public property let action ( prm_action ) : pub_objCPU.properties.item("action").value = prm_action : end property
	public property let limit ( prm_limit ) : pub_objCPU.properties.item("limit").value = prm_limit : end property
	public property let resetInterval ( prm_resetInterval ) : pub_objCPU.properties.item("resetInterval").value = prm_resetInterval : end property
	public property let smpAffinitized ( prm_smpAffinitized ) : pub_objCPU.properties.item("smpAffinitized").value = prm_smpAffinitized : end property
	public property let smpProcessorAffinityMask ( prm_smpProcessorAffinityMask ) : pub_objCPU.properties.item("smpProcessorAffinityMask").value = prm_smpProcessorAffinityMask : end property
	public property let smpProcessorAffinityMask2 ( prm_smpProcessorAffinityMask2 ) : pub_objCPU.properties.item("smpProcessorAffinityMask2").value = prm_smpProcessorAffinityMask2 : end property

	public pub_objCPU
end class

'*****************
'* Recycling
'*****************
class clsRecycling
	public default property get exists : exists = true : end property
	public property get disallowOverlappingRotation : disallowOverlappingRotation = pub_objRecycling.properties.item("disallowOverlappingRotation").value : end property
	public property get disallowRotationOnConfigChange : disallowRotationOnConfigChange = pub_objRecycling.properties.item("disallowRotationOnConfigChange").value : end property
	public property get logEventOnRecycle : logEventOnRecycle = pub_objRecycling.properties.item("logEventOnRecycle").value : end property

 	public property let disallowOverlappingRotation ( prm_disallowOverlappingRotation ) : pub_objRecycling.properties.item("disallowOverlappingRotation").value = prm_disallowOverlappingRotation : end property
	public property let disallowRotationOnConfigChange ( prm_disallowRotationOnConfigChange ) : pub_objRecycling.properties.item("disallowRotationOnConfigChange").value = prm_disallowRotationOnConfigChange : end property
	public property let logEventOnRecycle ( prm_logEventOnRecycle ) : pub_objRecycling.properties.item("logEventOnRecycle").value = prm_logEventOnRecycle : end property

   public property get periodicRestart
		dim pvt_objPeriodicRestart
		set pvt_objPeriodicRestart = new clsPeriodicRestart
		set pvt_objPeriodicRestart.pub_objPeriodicRestart = pub_objRecycling.childElements.item("periodicRestart")
		set periodicRestart = pvt_objPeriodicRestart
	end property

	public pub_objRecycling
end class

'********************
'* Periodic Restart
'********************
class clsPeriodicRestart
	public default property get exists : exists = true : end property
	public property get memory : memory = pub_objPeriodicRestart.properties.item("memory").value : end property
	public property get privateMemory : privateMemory = pub_objPeriodicRestart.properties.item("privateMemory").value : end property
	public property get requests : requests = pub_objPeriodicRestart.properties.item("requests").value : end property
	public property get time : time = pub_objPeriodicRestart.properties.item("time").value : end property

  	public property let memory ( prm_memory ) : pub_objPeriodicRestart.properties.item("memory").value = prm_memory : end property
	public property let privateMemory ( prm_privateMemory ) : pub_objPeriodicRestart.properties.item("privateMemory").value = prm_privateMemory : end property
	public property let requests ( prm_requests ) : pub_objPeriodicRestart.properties.item("requests").value = prm_requests : end property
	public property let time ( prm_time ) : pub_objPeriodicRestart.properties.item("time").value = prm_time : end property

	public pub_objPeriodicRestart
end class

'********************
'* Process Model
'********************
class clsProcessModel
	public default property get exists : exists = true : end property
	public property get identityType : identityType = pub_objProcessModel.properties.item("identityType").value : end property
	public property get idleTimeout : idleTimeout = pub_objProcessModel.properties.item("idleTimeout").value : end property
	public property get loadUserProfile : loadUserProfile = pub_objProcessModel.properties.item("loadUserProfile").value : end property
	public property get logonType : logonType = pub_objProcessModel.properties.item("logonType").value : end property
	public property get manualGroupMembership : manualGroupMembership = pub_objProcessModel.properties.item("manualGroupMembership").value : end property
	public property get maxProcesses : maxProcesses = pub_objProcessModel.properties.item("maxProcesses").value : end property
	public property get password : password = pub_objProcessModel.properties.item("password").value : end property
	public property get pingingEnabled : pingingEnabled = pub_objProcessModel.properties.item("pingingEnabled").value : end property
	public property get pingInterval : pingInterval = pub_objProcessModel.properties.item("pingInterval").value : end property
	public property get pingResponseTime : pingResponseTime = pub_objProcessModel.properties.item("pingResponseTime").value : end property
	public property get shutdownTimeLimit : shutdownTimeLimit = pub_objProcessModel.properties.item("shutdownTimeLimit").value : end property
	public property get startupTimeLimit : startupTimeLimit = pub_objProcessModel.properties.item("startupTimeLimit").value : end property
	public property get userName : userName = pub_objProcessModel.properties.item("userName").value : end property

  	public property let identityType ( prm_identityType ) : pub_objProcessModel.properties.item("identityType").value = prm_identityType : end property
	public property let idleTimeout ( prm_idleTimeout ) : pub_objProcessModel.properties.item("idleTimeout").value = prm_idleTimeout : end property
	public property let loadUserProfile ( prm_loadUserProfile ) : pub_objProcessModel.properties.item("loadUserProfile").value = prm_loadUserProfile : end property
	public property let logonType ( prm_logonType ) : pub_objProcessModel.properties.item("logonType").value = prm_logonType : end property
	public property let manualGroupMembership ( prm_manualGroupMembership ) : pub_objProcessModel.properties.item("manualGroupMembership").value = prm_manualGroupMembership : end property
	public property let maxProcesses ( prm_maxProcesses ) : pub_objProcessModel.properties.item("maxProcesses").value = prm_maxProcesses : end property
	public property let password ( prm_password ) : pub_objProcessModel.properties.item("password").value = prm_password : end property
	public property let pingingEnabled ( prm_pingingEnabled ) : pub_objProcessModel.properties.item("pingingEnabled").value = prm_pingingEnabled : end property
	public property let pingInterval ( prm_pingInterval ) : pub_objProcessModel.properties.item("pingInterval").value = prm_pingInterval : end property
	public property let pingResponseTime ( prm_pingResponseTime ) : pub_objProcessModel.properties.item("pingResponseTime").value = prm_pingResponseTime : end property
	public property let shutdownTimeLimit ( prm_shutdownTimeLimit ) : pub_objProcessModel.properties.item("shutdownTimeLimit").value = prm_shutdownTimeLimit : end property
	public property let startupTimeLimit ( prm_startupTimeLimit ) : pub_objProcessModel.properties.item("startupTimeLimit").value = prm_startupTimeLimit : end property
	public property let userName ( prm_userName ) : pub_objProcessModel.properties.item("userName").value = prm_userName : end property

	public pub_objProcessModel
end class

'***************
'* Failure
'***************
class clsFailure
	public default property get exists : exists = true : end property
	public property get autoShutdownExe : autoShutdownExe = pub_objFailure.properties.item("autoShutdownExe").value : end property
	public property get autoShutdownParams : autoShutdownParams = pub_objFailure.properties.item("autoShutdownParams").value : end property
	public property get loadBalancerCapabilities : loadBalancerCapabilities = pub_objFailure.properties.item("loadBalancerCapabilities").value : end property
	public property get orphanActionExe : orphanActionExe = pub_objFailure.properties.item("orphanActionExe").value : end property
	public property get orphanActionParams : orphanActionParams = pub_objFailure.properties.item("orphanActionParams").value : end property
	public property get orphanWorkerProcess : orphanWorkerProcess = pub_objFailure.properties.item("orphanWorkerProcess").value : end property
	public property get rapidFailProtection : rapidFailProtection = pub_objFailure.properties.item("rapidFailProtection").value : end property
	public property get rapidFailProtectionInterval : rapidFailProtectionInterval = pub_objFailure.properties.item("rapidFailProtectionInterval").value : end property
	public property get rapidFailProtectionMaxCrashes : rapidFailProtectionMaxCrashes = pub_objFailure.properties.item("rapidFailProtectionMaxCrashes").value : end property

	public property let autoShutdownExe ( prm_autoShutdownExe ) : pub_objFailure.properties.item("autoShutdownExe").value = prm_autoShutdownExe : end property
	public property let autoShutdownParams ( prm_autoShutdownParams ) : pub_objFailure.properties.item("autoShutdownParams").value = prm_autoShutdownParams : end property
	public property let loadBalancerCapabilities ( prm_loadBalancerCapabilities ) : pub_objFailure.properties.item("loadBalancerCapabilities").value = prm_loadBalancerCapabilities : end property
	public property let orphanActionExe ( prm_orphanActionExe ) : pub_objFailure.properties.item("orphanActionExe").value = prm_orphanActionExe : end property
	public property let orphanActionParams ( prm_orphanActionParams ) : pub_objFailure.properties.item("orphanActionParams").value = prm_orphanActionParams : end property
	public property let orphanWorkerProcess ( prm_orphanWorkerProcess ) : pub_objFailure.properties.item("orphanWorkerProcess").value = prm_orphanWorkerProcess : end property
	public property let rapidFailProtection ( prm_rapidFailProtection ) : pub_objFailure.properties.item("rapidFailProtection").value = prm_rapidFailProtection : end property
	public property let rapidFailProtectionInterval ( prm_rapidFailProtectionInterval ) : pub_objFailure.properties.item("rapidFailProtectionInterval").value = prm_rapidFailProtectionInterval : end property
	public property let rapidFailProtectionMaxCrashes ( prm_rapidFailProtectionMaxCrashes ) : pub_objFailure.properties.item("rapidFailProtectionMaxCrashes").value = prm_rapidFailProtectionMaxCrashes : end property

	public pub_objFailure
end class

'*****************
'* FTP Server
'*****************
class clsFtpServer
	public default property get exists : exists = true : end property
	public property get allowUTF8 : allowUTF8 = pub_objFtpServer.properties.item("allowUTF8").value : end property
	public property get lastStartupStatus : lastStartupStatus = pub_objFtpServer.properties.item("lastStartupStatus").value : end property
	public property get serverAutoStart : serverAutoStart = pub_objFtpServer.properties.item("serverAutoStart").value : end property

  	public property let allowUTF8 ( prm_allowUTF8 ) : pub_objFtpServer.properties.item("allowUTF8").value = prm_allowUTF8 : end property
	public property let lastStartupStatus ( prm_lastStartupStatus ) : pub_objFtpServer.properties.item("lastStartupStatus").value = prm_lastStartupStatus : end property
	public property let serverAutoStart ( prm_serverAutoStart ) : pub_objFtpServer.properties.item("serverAutoStart").value = prm_serverAutoStart : end property

   public property get connections
		dim pvt_objConnections
		set pvt_objConnections = new clsConnections
		set pvt_objConnections.pub_objConnections = pub_objFtpServer.childElements.item("connections")
		set connections = pvt_objConnections
	end property

   public property get directoryBrowse
		dim pvt_objDirectoryBrowse
		set pvt_objDirectoryBrowse = new clsDirectoryBrowse
		set pvt_objDirectoryBrowse.pub_objDirectoryBrowse = pub_objFtpServer.childElements.item("directoryBrowse")
		set directoryBrowse = pvt_objDirectoryBrowse
	end property

   public property get fileHandling
		dim pvt_objFileHandling
		set pvt_objFileHandling = new clsFileHandling
		set pvt_objFileHandling.pub_objFileHandling = pub_objFtpServer.childElements.item("fileHandling")
		set fileHandling = pvt_objFileHandling
	end property

   public property get firewallSupport
		dim pvt_objFirewallSupport
		set pvt_objFirewallSupport = new clsFirewallSupport
		set pvt_objFirewallSupport.pub_objFirewallSupport = pub_objFtpServer.childElements.item("firewallSupport")
		set firewallSupport = pvt_objFirewallSupport
	end property

   public property get logFile
		dim pvt_objLogFile
		set pvt_objLogFile = new clsLogFile
		set pvt_objLogFile.pub_objLogFile = pub_objFtpServer.childElements.item("logFile")
		set logFile = pvt_objLogFile
	end property

   public property get messages
		dim pvt_objMessages
		set pvt_objMessages = new clsMessages
		set pvt_objMessages.pub_objMessages = pub_objFtpServer.childElements.item("messages")
		set messages = pvt_objMessages
	end property

   public property get security
		dim pvt_objSecurity
		set pvt_objSecurity = new clsSecurity
		set pvt_objSecurity.pub_objSecurity = pub_objFtpServer.childElements.item("security")
		set security = pvt_objSecurity
	end property

   public property get userIsolation
		dim pvt_objUserIsolation
		set pvt_objUserIsolation = new clsUserIsolation
		set pvt_objUserIsolation.pub_objUserIsolation = pub_objFtpServer.childElements.item("userIsolation")
		set userIsolation = pvt_objUserIsolation
	end property

	public pub_objFtpServer
end class

'*******************
'* Connections
'*******************
class clsConnections
	public default property get exists : exists = true : end property
	public property get controlChannelTimeout : controlChannelTimeout = pub_objConnections.properties.item("controlChannelTimeout").value : end property
	public property get dataChannelTimeout : dataChannelTimeout = pub_objConnections.properties.item("dataChannelTimeout").value : end property
	public property get disableSocketPooling : disableSocketPooling = pub_objConnections.properties.item("disableSocketPooling").value : end property
	public property get maxBandwidth : maxBandwidth = pub_objConnections.properties.item("maxBandwidth").value : end property
	public property get maxConnections : maxConnections = pub_objConnections.properties.item("maxConnections").value : end property
	public property get minBytesPerSecond : minBytesPerSecond = pub_objConnections.properties.item("minBytesPerSecond").value : end property
	public property get resetOnMaxConnections : resetOnMaxConnections = pub_objConnections.properties.item("resetOnMaxConnections").value : end property
	public property get serverListenBacklog : serverListenBacklog = pub_objConnections.properties.item("serverListenBacklog").value : end property
	public property get unauthenticatedTimeout : unauthenticatedTimeout = pub_objConnections.properties.item("unauthenticatedTimeout").value : end property

  	public property let controlChannelTimeout ( prm_controlChannelTimeout ) : pub_objConnections.properties.item("controlChannelTimeout").value = prm_controlChannelTimeout : end property
	public property let dataChannelTimeout ( prm_dataChannelTimeout ) : pub_objConnections.properties.item("dataChannelTimeout").value = prm_dataChannelTimeout : end property
	public property let disableSocketPooling ( prm_disableSocketPooling ) : pub_objConnections.properties.item("disableSocketPooling").value = prm_disableSocketPooling : end property
	public property let maxBandwidth ( prm_maxBandwidth ) : pub_objConnections.properties.item("maxBandwidth").value = prm_maxBandwidth : end property
	public property let maxConnections ( prm_maxConnections ) : pub_objConnections.properties.item("maxConnections").value = prm_maxConnections : end property
	public property let minBytesPerSecond ( prm_minBytesPerSecond ) : pub_objConnections.properties.item("minBytesPerSecond").value = prm_minBytesPerSecond : end property
	public property let resetOnMaxConnections ( prm_resetOnMaxConnections ) : pub_objConnections.properties.item("resetOnMaxConnections").value = prm_resetOnMaxConnections : end property
	public property let serverListenBacklog ( prm_serverListenBacklog ) : pub_objConnections.properties.item("serverListenBacklog").value = prm_serverListenBacklog : end property
	public property let unauthenticatedTimeout ( prm_unauthenticatedTimeout ) : pub_objConnections.properties.item("unauthenticatedTimeout").value = prm_unauthenticatedTimeout : end property

	public pub_objConnections
end class

'***********************
'* Directory Browse
'***********************
class clsDirectoryBrowse
	public default property get exists : exists = true : end property
	public property get showFlags : showFlags = pub_objDirectoryBrowse.properties.item("showFlags").value : end property
	public property get virtualDirectoryTimeout : virtualDirectoryTimeout = pub_objDirectoryBrowse.properties.item("virtualDirectoryTimeout").value : end property

  	public property let showFlags ( prm_showFlags ) : pub_objDirectoryBrowse.properties.item("showFlags").value = prm_showFlags : end property
	public property let virtualDirectoryTimeout ( prm_virtualDirectoryTimeout ) : pub_objDirectoryBrowse.properties.item("virtualDirectoryTimeout").value = prm_virtualDirectoryTimeout : end property

	public pub_objDirectoryBrowse
end class

'********************
'* File Handling
'********************
class clsFileHandling
	public default property get exists : exists = true : end property
	public property get allowReadUploadsInProgress : allowReadUploadsInProgress = pub_objFileHandling.properties.item("allowReadUploadsInProgress").value : end property
	public property get allowReplaceOnRename : allowReplaceOnRename = pub_objFileHandling.properties.item("allowReplaceOnRename").value : end property

	public property let allowReadUploadsInProgress ( prm_allowReadUploadsInProgress ) : pub_objFileHandling.properties.item("allowReadUploadsInProgress").value = prm_allowReadUploadsInProgress : end property
	public property let allowReplaceOnRename ( prm_allowReplaceOnRename ) : pub_objFileHandling.properties.item("allowReplaceOnRename").value = prm_allowReplaceOnRename : end property

	public pub_objFileHandling
end class

'***********************
'* Firewall Support
'***********************
class clsFirewallSupport
	public default property get exists : exists = true : end property
	public property get externalIp4Address : externalIp4Address = pub_objFirewallSupport.properties.item("externalIp4Address").value : end property

  	public property let externalIp4Address ( prm_externalIp4Address ) : pub_objFirewallSupport.properties.item("externalIp4Address").value = prm_externalIp4Address : end property

	public pub_objFirewallSupport
end class

'***************
'* Log File
'***************
class clsLogFile
	public default property get exists : exists = true : end property
	public property get customLogPluginClsid : customLogPluginClsid = pub_objLogFile.properties.item("customLogPluginClsid").value : end property
	public property get directory : directory = pub_objLogFile.properties.item("directory").value : end property
	public property get enabled : enabled = pub_objLogFile.properties.item("enabled").value : end property
	public property get localTimeRollover : localTimeRollover = pub_objLogFile.properties.item("localTimeRollover").value : end property
	public property get logExtFileFlags : logExtFileFlags = pub_objLogFile.properties.item("logExtFileFlags").value : end property
	public property get logFormat : logFormat = pub_objLogFile.properties.item("logFormat").value : end property
	public property get period : period = pub_objLogFile.properties.item("period").value : end property
	public property get selectiveLogging : selectiveLogging = pub_objLogFile.properties.item("selectiveLogging").value : end property
	public property get truncateSize : truncateSize = pub_objLogFile.properties.item("truncateSize").value : end property

  	public property let customLogPluginClsid ( prm_customLogPluginClsid ) : pub_objLogFile.properties.item("customLogPluginClsid").value = prm_customLogPluginClsid : end property
	public property let directory ( prm_directory ) : pub_objLogFile.properties.item("directory").value = prm_directory : end property
	public property let enabled ( prm_enabled ) : pub_objLogFile.properties.item("enabled").value = prm_enabled : end property
	public property let localTimeRollover ( prm_localTimeRollover ) : pub_objLogFile.properties.item("localTimeRollover").value = prm_localTimeRollover : end property
	public property let logExtFileFlags ( prm_logExtFileFlags ) : pub_objLogFile.properties.item("logExtFileFlags").value = prm_logExtFileFlags : end property
	public property let logFormat ( prm_logFormat ) : pub_objLogFile.properties.item("logFormat").value = prm_logFormat : end property
	public property let period ( prm_period ) : pub_objLogFile.properties.item("period").value = prm_period : end property
	public property let selectiveLogging ( prm_selectiveLogging ) : pub_objLogFile.properties.item("selectiveLogging").value = prm_selectiveLogging : end property
	public property let truncateSize ( prm_truncateSize ) : pub_objLogFile.properties.item("truncateSize").value = prm_truncateSize : end property

	public pub_objLogFile
end class

'**********
'* Messages
'**********
class clsMessages
	public default property get exists : exists = true : end property
	public property get allowLocalDetailedErrors : allowLocalDetailedErrors = pub_objMessages.properties.item("allowLocalDetailedErrors").value : end property
	public property get bannerMessage : bannerMessage = pub_objMessages.properties.item("bannerMessage").value : end property
	public property get exitMessage : exitMessage = pub_objMessages.properties.item("exitMessage").value : end property
	public property get expandVariables : expandVariables = pub_objMessages.properties.item("expandVariables").value : end property
	public property get greetingMessage : greetingMessage = pub_objMessages.properties.item("greetingMessage").value : end property
	public property get maxClientsMessage : maxClientsMessage = pub_objMessages.properties.item("maxClientsMessage").value : end property
	public property get suppressDefaultBanner : suppressDefaultBanner = pub_objMessages.properties.item("suppressDefaultBanner").value : end property

  	public property let allowLocalDetailedErrors ( prm_allowLocalDetailedErrors ) : pub_objMessages.properties.item("allowLocalDetailedErrors").value = prm_allowLocalDetailedErrors : end property
	public property let bannerMessage ( prm_bannerMessage ) : pub_objMessages.properties.item("bannerMessage").value = prm_bannerMessage : end property
	public property let exitMessage ( prm_exitMessage ) : pub_objMessages.properties.item("exitMessage").value = prm_exitMessage : end property
	public property let expandVariables ( prm_expandVariables ) : pub_objMessages.properties.item("expandVariables").value = prm_expandVariables : end property
	public property let greetingMessage ( prm_greetingMessage ) : pub_objMessages.properties.item("greetingMessage").value = prm_greetingMessage : end property
	public property let maxClientsMessage ( prm_maxClientsMessage ) : pub_objMessages.properties.item("maxClientsMessage").value = prm_maxClientsMessage : end property
	public property let suppressDefaultBanner ( prm_suppressDefaultBanner ) : pub_objMessages.properties.item("suppressDefaultBanner").value = prm_suppressDefaultBanner : end property

	public pub_objMessages
end class

'****************
'* Security
'****************
class clsSecurity
	public default property get exists : exists = true : end property
   public property get authentication
		dim pvt_objAuthentication
		set pvt_objAuthentication = new clsAuthentication
		set pvt_objAuthentication.pub_objAuthentication = pub_objSecurity.childElements.item("authentication")
		set authentication = pvt_objAuthentication
	end property

   public property get commandFiltering
		dim pvt_objCommandFiltering
		set pvt_objCommandFiltering = new clsCommandFiltering
		set pvt_objCommandFiltering.pub_objCommandFiltering = pub_objSecurity.childElements.item("commandFiltering")
		set commandFiltering = pvt_objCommandFiltering
	end property

   public property get dataChannelSecurity
		dim pvt_objDataChannelSecurity
		set pvt_objDataChannelSecurity = new clsDataChannelSecurity
		set pvt_objDataChannelSecurity.pub_objDataChannelSecurity = pub_objSecurity.childElements.item("dataChannelSecurity")
		set dataChannelSecurity = pvt_objDataChannelSecurity
	end property

   public property get ssl
		dim pvt_objSsl
		set pvt_objSsl = new clsSsl
		set pvt_objSsl.pub_objSsl = pub_objSecurity.childElements.item("ssl")
		set ssl = pvt_objSsl
	end property

   public property get sslClientCertificates
		dim pvt_objSslClientCertificates
		set pvt_objSslClientCertificates = new clsSslClientCertificates
		set pvt_objSslClientCertificates.pub_objSslClientCertificates = pub_objSecurity.childElements.item("sslClientCertificates")
		set sslClientCertificates = pvt_objSslClientCertificates
	end property

	public pub_objSecurity
end class

'*************************
'* SSL Client Certificates
'*************************
class clsSslClientCertificates
	public default property get exists : exists = true : end property
	public property get clientCertificatePolicy : clientCertificatePolicy = pub_objSslClientCertificates.properties.item("clientCertificatePolicy").value : end property
	public property get revocationFreshnessTime : revocationFreshnessTime = pub_objSslClientCertificates.properties.item("revocationFreshnessTime").value : end property
	public property get revocationURLRetrievalTimeout : revocationURLRetrievalTimeout = pub_objSslClientCertificates.properties.item("revocationURLRetrievalTimeout").value : end property
	public property get useActiveDirectoryMapping : useActiveDirectoryMapping = pub_objSslClientCertificates.properties.item("useActiveDirectoryMapping").value : end property
	public property get validationFlags : validationFlags = pub_objSslClientCertificates.properties.item("validationFlags").value : end property

	public property let clientCertificatePolicy ( prm_clientCertificatePolicy ) : pub_objSslClientCertificates.properties.item("clientCertificatePolicy").value = prm_clientCertificatePolicy : end property
	public property let revocationFreshnessTime ( prm_revocationFreshnessTime ) : pub_objSslClientCertificates.properties.item("revocationFreshnessTime").value = prm_revocationFreshnessTime : end property
	public property let revocationURLRetrievalTimeout ( prm_revocationURLRetrievalTimeout ) : pub_objSslClientCertificates.properties.item("revocationURLRetrievalTimeout").value = prm_revocationURLRetrievalTimeout : end property
	public property let useActiveDirectoryMapping ( prm_useActiveDirectoryMapping ) : pub_objSslClientCertificates.properties.item("useActiveDirectoryMapping").value = prm_useActiveDirectoryMapping : end property
	public property let validationFlags ( prm_validationFlags ) : pub_objSslClientCertificates.properties.item("validationFlags").value = prm_validationFlags : end property

	public pub_objSslClientCertificates
end class

'***********
'* SSL
'***********
class clsSsl
	public default property get exists : exists = true : end property
	public property get controlChannelPolicy : controlChannelPolicy = pub_objSsl.properties.item("controlChannelPolicy").value : end property
	public property get dataChannelPolicy : dataChannelPolicy = pub_objSsl.properties.item("dataChannelPolicy").value : end property
	public property get serverCertHash : serverCertHash = pub_objSsl.properties.item("serverCertHash").value : end property
	public property get serverCertStoreName : serverCertStoreName = pub_objSsl.properties.item("serverCertStoreName").value : end property
	public property get ssl128 : ssl128 = pub_objSsl.properties.item("ssl128").value : end property

  	public property let controlChannelPolicy ( prm_controlChannelPolicy ) : pub_objSsl.properties.item("controlChannelPolicy").value = prm_controlChannelPolicy : end property
	public property let dataChannelPolicy ( prm_dataChannelPolicy ) : pub_objSsl.properties.item("dataChannelPolicy").value = prm_dataChannelPolicy : end property
	public property let serverCertHash ( prm_serverCertHash ) : pub_objSsl.properties.item("serverCertHash").value = prm_serverCertHash : end property
	public property let serverCertStoreName ( prm_serverCertStoreName ) : pub_objSsl.properties.item("serverCertStoreName").value = prm_serverCertStoreName : end property
	public property let ssl128 ( prm_ssl128 ) : pub_objSsl.properties.item("ssl128").value = prm_ssl128 : end property

	public pub_objSsl
end class

'***************************
'* Data Channel Security
'***************************
class clsDataChannelSecurity
	public default property get exists : exists = true : end property
	public property get matchClientAddressForPasv : matchClientAddressForPasv = pub_objDataChannelSecurity.properties.item("matchClientAddressForPasv").value : end property
	public property get matchClientAddressForPort : matchClientAddressForPort = pub_objDataChannelSecurity.properties.item("matchClientAddressForPort").value : end property

	public property let matchClientAddressForPasv ( prm_matchClientAddressForPasv ) : pub_objDataChannelSecurity.properties.item("matchClientAddressForPasv").value = prm_matchClientAddressForPasv : end property
	public property let matchClientAddressForPort ( prm_matchClientAddressForPort ) : pub_objDataChannelSecurity.properties.item("matchClientAddressForPort").value = prm_matchClientAddressForPort : end property

	public pub_objDataChannelSecurity
end class

'************************
'* Command Filtering
'************************
class clsCommandFiltering
	public default property get exists : exists = true : end property
	public property get allowUnlisted : allowUnlisted = pub_objCommandFiltering.properties.item("allowUnlisted").value : end property
	public property get maxCommandLine : maxCommandLine = pub_objCommandFiltering.properties.item("maxCommandLine").value : end property

  	public property let allowUnlisted ( prm_allowUnlisted ) : pub_objCommandFiltering.properties.item("allowUnlisted").value = prm_allowUnlisted : end property
	public property let maxCommandLine ( prm_maxCommandLine ) : pub_objCommandFiltering.properties.item("maxCommandLine").value = prm_maxCommandLine : end property

	public pub_objCommandFiltering
end class

'**********************
'* Authentication
'**********************
class clsAuthentication
	public default property get exists : exists = true : end property
   public property get anonymousAuthentication
		dim pvt_objAnonymousAuthentication
		set pvt_objAnonymousAuthentication = new clsAnonymousAuthentication
		set pvt_objAnonymousAuthentication.pub_objAnonymousAuthentication = pub_objAuthentication.childElements.item("anonymousAuthentication")
		set anonymousAuthentication = pvt_objAnonymousAuthentication
	end property

   public property get basicAuthentication
		dim pvt_objBasicAuthentication
		set pvt_objBasicAuthentication = new clsBasicAuthentication
		set pvt_objBasicAuthentication.pub_objBasicAuthentication = pub_objAuthentication.childElements.item("basicAuthentication")
		set basicAuthentication = pvt_objBasicAuthentication
	end property

   public property get clientCertAuthentication
		dim pvt_objClientCertAuthentication
		set pvt_objClientCertAuthentication = new clsClientCertAuthentication
		set pvt_objClientCertAuthentication.pub_objClientCertAuthentication = pub_objAuthentication.childElements.item("clientCertAuthentication")
		set clientCertAuthentication = pvt_objClientCertAuthentication
	end property

	public pub_objAuthentication
end class

'*******************************
'* Anonymous Authentication
'*******************************
class clsAnonymousAuthentication
	public default property get exists : exists = true : end property
 	public property get defaultLogonDomain : defaultLogonDomain = pub_objAnonymousAuthentication.properties.item("defaultLogonDomain").value : end property
	public property get enabled : enabled = pub_objAnonymousAuthentication.properties.item("enabled").value : end property
	public property get logonMethod : logonMethod = pub_objAnonymousAuthentication.properties.item("logonMethod").value : end property
	public property get password : password = pub_objAnonymousAuthentication.properties.item("password").value : end property
	public property get userName : userName = pub_objAnonymousAuthentication.properties.item("userName").value : end property

 	public property let defaultLogonDomain ( prm_defaultLogonDomain ) : pub_objAnonymousAuthentication.properties.item("defaultLogonDomain").value = prm_defaultLogonDomain : end property
	public property let enabled ( prm_enabled ) : pub_objAnonymousAuthentication.properties.item("enabled").value = prm_enabled : end property
	public property let logonMethod ( prm_logonMethod ) : pub_objAnonymousAuthentication.properties.item("logonMethod").value = prm_logonMethod : end property
	public property let password ( prm_password ) : pub_objAnonymousAuthentication.properties.item("password").value = prm_password : end property
	public property let userName ( prm_userName ) : pub_objAnonymousAuthentication.properties.item("userName").value = prm_userName : end property

	public pub_objAnonymousAuthentication
end class

'**************************
'* Basic Authentication
'*********************
class clsBasicAuthentication
	public default property get exists : exists = true : end property
	public property get defaultLogonDomain : defaultLogonDomain = pub_objBasicAuthentication.properties.item("defaultLogonDomain").value : end property
	public property get enabled : enabled = pub_objBasicAuthentication.properties.item("enabled").value : end property
	public property get logonMethod : logonMethod = pub_objBasicAuthentication.properties.item("logonMethod").value : end property

	public property let defaultLogonDomain ( prm_defaultLogonDomain ) : pub_objBasicAuthentication.properties.item("defaultLogonDomain").value = prm_defaultLogonDomain : end property
	public property let enabled ( prm_enabled ) : pub_objBasicAuthentication.properties.item("enabled").value = prm_enabled : end property
	public property let logonMethod ( prm_logonMethod ) : pub_objBasicAuthentication.properties.item("logonMethod").value = prm_logonMethod : end property

	public pub_objBasicAuthentication
end class

'**************************
'* Client Cert Authentication
'*********************
class clsClientCertAuthentication
	public default property get exists : exists = true : end property
	public property get enabled : enabled = pub_objClientCertAuthentication.properties.item("enabled").value : end property

  	public property let enabled ( prm_enabled ) : pub_objClientCertAuthentication.properties.item("enabled").value = prm_enabled : end property

	public pub_objClientCertAuthentication
end class

'*********************
'* User Isolation
'*********************
class clsUserIsolation
	public default property get exists : exists = true : end property
	public property get mode : mode = pub_objUserIsolation.properties.item("mode").value : end property

  	public property let mode ( prm_mode ) : pub_objUserIsolation.properties.item("mode").value = prm_mode : end property

   public property get activeDirectory
		dim pvt_objActiveDirectory
		set pvt_objActiveDirectory = new clsActiveDirectory
		set pvt_objActiveDirectory.pub_objActiveDirectory = pub_objUserIsolation.childElements.item("activeDirectory")
		set activeDirectory = pvt_objActiveDirectory
	end property

	public pub_objUserIsolation
end class

'*********************
'* Active Directory
'*********************
class clsActiveDirectory
	public default property get exists : exists = true : end property
	public property get adCacheRefresh : adCacheRefresh = pub_objActiveDirectory.properties.item("adCacheRefresh").value : end property
	public property get adPassword : adPassword = pub_objActiveDirectory.properties.item("adPassword").value : end property
	public property get adUserName : adUserName = pub_objActiveDirectory.properties.item("adUserName").value : end property

  	public property let adCacheRefresh ( prm_adCacheRefresh ) : pub_objActiveDirectory.properties.item("adCacheRefresh").value = prm_adCacheRefresh : end property
	public property let adPassword ( prm_adPassword ) : pub_objActiveDirectory.properties.item("adPassword").value = prm_adPassword : end property
	public property let adUserName ( prm_adUserName ) : pub_objActiveDirectory.properties.item("adUserName").value = prm_adUserName : end property

	public pub_objActiveDirectory
end class

'*******************
'* Utility Functions
'*******************
function intFindCollectionMemberByPropertyValue ( colCollection, strPropertyName, varPropertyValue )
	dim intMatchIndex
	dim intLoop
	intMatchIndex = -1
	for intLoop = 0 to (cInt(colCollection.count) - 1)
		if isObject(colCollection.item(intLoop).properties.item(cStr(strPropertyName))) then
			if strComp(colCollection.item(intLoop).properties.item(cStr(strPropertyName)).value,cStr(varPropertyValue),vbTextCompare) = 0 then
				intMatchIndex = intLoop
			end if
		end if
	next
	intFindCollectionMemberByPropertyValue = intMatchIndex
end Function