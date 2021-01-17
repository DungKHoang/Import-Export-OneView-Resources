# ------------------ Parameters
Param (
		[string]$workBook,
		[string]$scriptFolder=$pwd.Path
)


$DoubleQuote    = '"'
$CRLF           = "`r`n"
$Delimiter      = "\"   # Delimiter for CSV profile file
$SepHash        = ";"   # USe for multiple values fields
$hash           = '@'
$SepChar        = '|'
$CRLF           = "`r`n"
$OpenDelim      = "{"
$CloseDelim     = "}" 
$OpenArray      = "("
$CloseArray     = ")"
$CR             = "`n"
$Comma          = ','
$Equal          = '='
$Dot            = '.'
$Underscore     = '_'
$HexPattern     = "^[0-9a-fA-F][0-9a-fA-F]:"
$Space          = ' '
$TAB 			= '    '

$Syn12K         = 'SY12000' # Synergy enclosure type
$MAXLEN 		= 42


#------------------- Interconnect Types
$ICTypes         = @{
    "571956-B21" =  "FlexFabric" ;
    "455880-B21" =  "Flex10"     ;
    "638526-B21" =  "Flex1010D"  ;
    "691367-B21" =  "Flex2040f8" ;
    "572018-B21" =  "VCFC20"     ;
    "466482-B21" =  "VCFC24"     ;
    "641146-B21" =  "FEX"
}

# ---------------------- FC networks
[HashTable]$Global:FCNetworkFabricTypeEnum = @{

        FA           = 'FabricAttach';
        FabricAttach = 'FabricAttach';
        DA           = 'DirectAttach';
        DirectAttach = 'DirectAttach'

    }
    [Hashtable]$global:GetUplinkSetPortSpeeds    = @{
        Speed0M   = "0";
        Speed100M = "100Mb";
        Speed10G  = "10Gb";
        Speed10M  = "10Mb";
        Speed1G   = "1Gb";
        Speed1M   = "1Mb";
        Speed20G  = "20Gb";
        Speed2G   = "2Gb";
        Speed2_5G = "2.5Gb";
        Speed40G  = "40Gb";
        Speed4G   = "4Gb";
        Speed8G   = "8Gb";
        Auto       = "Auto"
    }
    [Hashtable]$global:SetUplinkSetPortSpeeds    = @{
        '0'    = "Speed0M";
        '100M' = "Speed100M";
        '100'  = "Speed100M";
        '10G'  = "Speed10G";
        '10'   = "Speed10G";
        '10M'  = "Speed10M";
        '1G'   = "Speed1G";
        '1'    = "Speed1G";
        '1M'   = "Speed1M";
        '20G'  = "Speed20G";
        '2G'   = "Speed2G";
        '2'    = "Speed2G";
        '2.5G' = "Speed2_5G";
        '40G'  = "Speed40G";
        '4G'   = "Speed4G";
        '8G'   = "Speed8G";
        '4'    = "Speed4G";
        '8'    = "Speed8G";
        'Auto' = "Auto"
    }

#-------------------- Resource type

$ResourceCategoryEnum = @{
    Baseline                    = 'firmware-drivers';
    ServerHardware              = 'server-hardware';
    ServerHardwareType          = 'server-hardware-types';
    ServerProfile               = 'server-profiles';
    ServerProfileTemplate       = 'server-profile-templates';
    Enclosure                   = 'enclosures';
    LogicalEnclosure            = 'logical-enclosures';
    EnclosureGroup              = 'enclosure-groups';
    Interconnect                = 'interconnects';
    LogicalInterconnect         = 'logical-interconnects';
    LogicalInterconnectGroup    = 'logical-interconnect-groups';
    LogicalSwitch               = 'logical-switches';
    LogicalSwitchGroup          = 'logical-switch-groups';
    UplinkSet                   = 'uplink-sets';
    SasInterconnect             = 'sas-interconnects';
    SasLogicalInterconnect      = 'sas-logical-interconnects';
    SasLogicalInterconnectGroup = 'sas-logical-interconnect-groups';
    ClusterProfile              = 'hypervisor-cluster-profiles';
    ClusterNode                 = 'hypervisor-host-profiles';
    HypervisorManager           = 'hypervisor-managers';
    HypervisorCluster           = 'hypervisor-cluster-profiles';
    FabricManager               = 'fabric-managers';
    FabricManagerTenant         = 'tenants';
    RackManager                 = 'rack-managers';
    NetworkSet                  = 'network-sets';
    EthernetNetwork             = 'ethernet-networks';
    StorageVolumeSet            = 'storage-volume-sets';
    StorageVolume               = 'storage-volumes';
    StorageVolumeTemplate       = 'storage-volume-templates';
    StoragePool                 = 'storage-pools';
    IPv4Subnet                  = 'id-range-IPv4-subnet';
    IPv4Range                   = 'id-range-IPv4';
    IPv6Subnet                  = 'id-range-IPv6-subnet';
    IPv6Range                   = 'id-range-IPv6';
    LogicalJBOD                 = 'sas-logical-jbods';
    Drive                       = 'drives';
    DriveEnclosure              = 'drive-enclosures'
}

#------------------- Interconnect Types
[Hashtable]$ICModuleTypes      = $ListofICTypes = @{
	"VirtualConnectSE40GbF8ModuleforSynergy"    =  "SEVC40f8";
	"Synergy20GbInterconnectLinkModule"         =  "SE20ILM";
	"Synergy10GbInterconnectLinkModule"         =  "SE10ILM";
	"VirtualConnectSE16GbFCModuleforSynergy"    =  "SEVC16GbFC";
	"Synergy12GbSASConnectionModule"            =  "SE12SAS"
}

[Hashtable]$FabricModuleTypes  = @{
	"VirtualConnectSE40GbF8ModuleforSynergy"    =  "SEVC40f8";
	"Synergy12GbSASConnectionModule"            =  "SAS";
	"VirtualConnectSE16GbFCModuleforSynergy"    =  "SEVCFC";
}

[Hashtable]$ServerProfileConnectionBootPriorityEnum = @{
	none           = 'NotBootable';
	NotBootable    = 'NotBootable';
	Primary        = 'Primary';
	Secondary      = 'Secondary';
	IscsiPrimary   = 'Primary';
	IscsiSecondary = 'Secondary';
	LoadBalanced   = 'LoadBalanced'
}
[Hashtable]$ServerProfileSanManageOSType            = @{
	CitrixXen  = "Citrix Xen Server 5.x/6.x";
	CitrisXen7 = "Citrix Xen Server 7.x";
	AIX        = "AIX";
	IBMVIO     = "IBM VIO Server";
	RHEL4      = "RHE Linux (Pre RHEL 5)";
	RHEL3      = "RHE Linux (Pre RHEL 5)";
	RHEL       = "RHE Linux (5.x, 6.x, 7.x)";
	RHEV       = "RHE Virtualization (5.x, 6.x)";
	RHEV7      = "RHE Virtualization 7.x";
	VMware     = "VMware (ESXi)";
	Win2k3     = "Windows 2003";
	Win2k8     = "Windows 2008/2008 R2";
	Win2k12    = "Windows 2012 / WS2012 R2";
	Win2k16    = "Windows Server 2016";
	OpenVMS    = "OpenVMS";
	Egenera    = "Egenera";
	Exanet     = "Exanet";
	Solaris9   = "Solaris 9/10";
	Solaris10  = "Solaris 9/10";
	Solaris11  = "Solaris 11";
	ONTAP      = "NetApp/ONTAP";
	OEL        = "OE Linux UEK (5.x, 6.x)";
	OEL7       = "OE Linux UEK 7.x";
	HPUX11iv1  = "HP-UX (11i v1, 11i v2)"
	HPUX11iv2  = "HP-UX (11i v1, 11i v2)";
	HPUX11iv3  = "HP-UX (11i v3)";
	SUSE       = "SuSE (10.x, 11.x, 12.x)";
	SUSE9      = "SuSE Linux (Pre SLES 10)";
	Inform     = "InForm"
}

[Hashtable]$SnmpAuthLevelEnum                    = @{
	None        = "noauthnopriv";
	AuthOnly    = "authnopriv";
	AuthAndPriv = "authpriv"
}
[Hashtable]$Snmpv3UserAuthLevelEnum              = @{
	None        = "None";
	AuthOnly    = "Authentication";
	AuthAndPriv = "Authentication and privacy"
}
[Hashtable]$SnmpAuthProtocolEnum                 = @{

	'none'   = 'none';
	'md5'    = 'MD5';
	'SHA'    = 'SHA';
	'sha1'   = 'SHA1';
	'sha2'   = 'SHA2';
	'sha256' = 'SHA256';
	'sha384' = 'SHA384';
	'sha512' = 'SHA512'

}
[Hashtable]$SnmpPrivProtocolEnum                 = @{
	'none'    = 'none';
	'aes'     = "AES128";
	'aes-128' = "AES128";
	'aes-192' = "AES192";
	'aes-256' = "AES256";
	'aes128'  = "AES128";
	'aes192'  = "AES192";
	'aes256'  = "AES256";
	'des56'   = "DES56";
	'3des'    = "3DES";
	'tdea'    = 'TDEA'
}
[Hashtable]$ApplianceSnmpV3PrivProtocolEnum      = @{
	'none'   = 'none';
	"des56"  = 'DES';
	'3des'   = '3DES';
	'aes128' = 'AES-128';
	'aes192' = 'AES-192';
	'aes256' = 'AES-256'
}

[Hashtable]$consistencyCheckingEnum 			= @{
	'NotChecked'	= 'None';
	'Not Checked'	= 'None';
	'UnChecked'		= 'None';
	''				= 'None';
	'None'			= 'None';

	'Checked'		= 'Exact';
	'ExactMatch'	= 'Exact';
	'Exact'			= 'Exact';

	'ExactMinimum'	= 'Minimum';
	'Minimum'		= 'Minimum'
	}

[HashTable]$iLOPrivilgeParamEnum     = @{
	'userConfigPriv'			= ' -AdministerUserAccounts $True '  ;
	'remoteConsolePriv'			= ' -RemoteConsole $True '			 ;
	'virtualMediaPriv'			= ' -VirtualMedia $True '			 ;
	'virtualPowerAndResetPriv'	= ' -VirtualPowerAndReset $True '	 ;
	'iloConfigPriv'				= ' -ConfigureIloSettings $True '	 ;
	'loginPriv'					= ' -Login $True '					 ;
	'hostBIOSConfigPriv'		= ' -HostBIOS $True '				 ;
	'hostNICConfigPriv'			= ' -HostNIC $True '				 ;
	'hostStorageConfigPriv'		= ' -HostStorage $True '			 
}
# ----------------------------------------------------
#
# 		OneView configuration
#
# ----------------------------------------------------

class OVnetwork
{
	[string]$hostName	
	[string]$domainName	
	[string]$ipV4	
	[string]$app1Ipv4	
	[string]$app2IpV4	
	[string]$ipv4Subnet	
	[string]$ipv4Gateway	
	[string]$ipv4Dns	
	[string]$ipV6	
	[string]$app1Ipv6	
	[string]$app2Ipv6	
	[string]$ipv6Subnet	
	[string]$ipv6Gateway	
	[string]$ipv6Dns	

}

class OVsec
{
	[string]$TLSname	
	[string]$mode	
	[string]$modeIsEnabled	
	[string]$enabled	
	[string]$cipherSuites	
}	

class OVAuth
{
	[string]$enable2FactorAuthentication	
	[string]$strictEnforcement	
	[string]$allowLocalLogin	
	[string]$allowEmergencyLogin	
	[string]$emergencyLoginType	
	[string]$directoryDomain	
	[string]$directoryDomainType	
	[string]$smartCardLoginOnly	
	[string]$validationOIDs							

}

# ----------------------------------------------------
#
# 		OneView settings
#
# ----------------------------------------------------

Class BackupConfig
{
	[boolean]$enabled
	[string]$remoteServerName
	[string]$remoteServerDir
	[string]$protocol
	[string]$port
	[string]$userName
	[string]$password 		= '***REDACTED***'
	[string]$scheduleInterval					# DAILY - WEEKLY
	[string]$scheduleDays						# MON|FRI|WED
	[string]$scheduleTime
	[string]$remoteServerPublicKey


}

Class Scope
{
	[string]$Name
	[string]$Description
	[string]$ResourceName
	[string]$ResourceType
}

Class ApplianceTimeLocale
{	
	[string]$locale          
	[string]$timezone        
	[string]$ntpServers      
	[string]$pollingInterval 
	[string]$syncWithHost
}

Class SmtpConfig
{             
	[string]$senderEmailAddress 
	[string]$password 		= '***REDACTED***'          
	[string]$smtpServer         
	[string]$smtpPort           
	[string]$smtpProtocol       
	[string]$alertEmailDisabled
	[string]$alertEmailFilters  
}

class firmwareBundle 
{
	[string]$name
	[string]$isofile
}

class repository
{
	[string]$name
	[string]$repositoryUrl
	[string]$directory
	[string]$username
	[string]$password 		= '***REDACTED***'
}

class proxy
{
	[string]$server
	[string]$port
	[string]$protocol
	[string]$username
}

class addressPool
{
	[string]$name
	[string]$poolType
	[boolean]$enabled
	[string]$rangeCategory
	[string]$startAddress
	[string]$endAddress
	[string]$networkId
	[string]$subnetmask
	[string]$gateway
	[string]$dnsServers
	[string]$domain
}

# ----------------------------------------------------
#
# 		OneView resources
#
# ----------------------------------------------------
class storageSystem
{
	[string]$name
	[string]$hostName
	[string]$familyName
	[string]$userName
	[string]$password		= '***REDACTED***'	
	[string]$systemPorts
	[string]$domainName
	[string]$vips
	[string]$model
	[string]$serialNumber
	[string]$wwnn
	[string]$firmware
}

class storagePool
{
	[string]$name	
	[string]$description	
	[string]$storageSystem	
	[string]$storageDomain	
	[string]$state	
	[string]$totalCapacity	
	[string]$allocatedCapacity	
	[string]$freeCapacity	
	[string]$driveType	
	[string]$RAID				
}

class storageVolumeTemplate
{
	[string]$name	
	[string]$description	
	[string]$familyName
	[string]$storageSystem	
	[string]$storagePool
	[string]$lockStoragePool
	[string]$snapshotStoragePool
	[string]$lockSnapshotStoragePool
	[string]$capacity
	[string]$lockCapacity
	[string]$provisioningType
	[string]$lockProvisioningType
	[string]$enableCompression	
	[string]$lockEnableCompression
	[string]$shared	
	[string]$lockProvisionMode
	[string]$enableAdaptiveOptimization	
	[string]$lockAdaptiveOptimization
	[string]$dataProtectionLevel
	[string]$lockDataProtectionLevel
	[string]$enableDeduplication	
	[string]$lockEnableDeduplication	
	[string]$lockPerformancePolicy
	[string]$enableEncryption		
	[string]$lockEnableEncryption
	[string]$cachePinning	
	[string]$lockCachePinning
	[string]$volumeSet
	[string]$lockVolumeSet
	[string]$enableIOPSLimit		
	[string]$lockEnableIOPSLimit
	[string]$enableDataTransferLimit		
	[string]$lockDataTransferLimit
	[string]$folder
	[string]$lockFolder
	[string]$scopes

}

class storageVolume
{
	[string]$name	
	[string]$description
	[string]$volumeTemplate
	[string]$storageSystem	
	[string]$storagePool
	[string]$snapshotStoragePool
	[string]$capacity
	[string]$provisioningType
	[string]$usedBy							
	[string]$enableCompression	
	[string]$shared	
	[string]$enableAdaptiveOptimization	
	[string]$dataProtectionLevel
	[string]$enableDeduplication	
	[string]$enableEncryption		
	[string]$cachePinning	
	[string]$volumeSet
	[string]$enableIOPSLimit		
	[string]$enableDataTransferLimit		
	[string]$folder
	[string]$scopes

}

class LogicalJBOD
{
	[string]$name	
	[string]$description	
	[string]$driveType	
	[int]$NumberofDrives	
	[string]$minDriveSize	
	[string]$maxDriveSize	
	[string]$driveEnclosure
	[string]$eraseDataonDelete	
	[string]$scopes
}
class ethernetNetwork 					
{
	[string]$name
	[string]$type
	[string]$vlanId   
	[string]$ethernetNetworkType
	[string]$subnetID	
	[string]$ipV6subnetID
	[string]$typicalBandwidth
	[string]$maximumBandwidth
	[string]$smartLink
	[string]$privateNetwork
	[string]$purpose  
	[string]$scopes
}

class fcFcoeNetwork 					
{
	[string]$name
	[string]$type
	[string]$fabricType
	[string]$managedSan
	[string]$vlanId   
	[string]$typicalBandwidth
	[string]$maximumBandwidth
	[string]$autoLoginRedistribution
	[string]$linkStabilityTime
	[string]$scopes
}

class networkSet 					
{
	[string]$name
	[string]$typicalBandwidth
	[string]$maximumBandwidth
	[string]$networkSetType
	[string]$networks
	[string]$nativeNetwork
	[string]$scopes
}

class lig
{
	[string]$name      
	[string]$frameCount
	[string]$interconnectBaySet
	[string]$enclosureType
	[string]$fabricModuleType
	[string]$bayConfig
	[string]$redundancyType
	[string]$internalNetworks
	[string]$consistencyCheckingForInternalNetworks
	[string]$interconnectConsistencyChecking
	[string]$enableIgmpSnooping
	[string]$igmpIdleTimeoutInterval
	[string]$enableFastMacCacheFailover
	[string]$macRefreshInterval
	[string]$enableNetworkLoopProtection
	[string]$enablePauseFloodProtection
	[string]$enableRichTLV 
	[string]$enableTaggedLldp
	[string]$lldpIpAddressMode
	[string]$lldpIpv4Address
	[string]$lldpIpv6Address
	[string]$enableStormControl
	[string]$stormControlPollingInterval
	[string]$stormControlThreshold
	[string]$qosconfigType
	[string]$downlinkClassificationType
	[string]$uplinkClassificationType
	[string]$scopes
	[string]$snmpConsistencyChecking


}


class uplinkset
{
	[string]$ligName
	[string]$name  
	[string]$networkType
	[string]$networks
	[string]$networkSets
	[string]$nativeNetwork
	[Boolean]$Trunking 
	[string]$logicalPortConfigInfos
	[string]$fcUplinkSpeed 
	[string]$loadBalancingMode
	[string]$lacpTimer
	[string]$primaryPort
	[string]$privateVLanDomains
	[string]$consistencyChecking
} 


# -- snmp 
class snmpConfiguration
{
	[string]$source						= 'Appliance'			# either 'Appliance'or lig_name
	[string]$format	
	[string]$communityString	
	[string]$contact	
	[string]$accessList	
	[string]$engineId
}

class snmpV3User
{
	#[boolean]$applianceSnmpUser		= $True 
	[string]$source 					= 'Appliance'			# either 'Appliance'or lig_name
	[string]$userName 					
	[string]$securityLevel 										# None - AuthOnly - AuthAndPriv
	[string]$authProtocol										# None - MD5 - SHA - SHA1 - SHA256 - SHA384 - SHA512
	[string]$authPassword 				=  '***REDACTED***'
	[string]$privacyProtocol									# None - des56 -3des - aes128 - aes192 -aes 256
	[string]$privacyPassword			=  '***REDACTED***'
}


class snmpTrap
{ 		
	[string]$source 					= 'Appliance'			# either 'Appliance'or lig_name	
	[string]$format						= 'SnmpV3'				# SnmpV1 or SnmpV3	
	[string]$destinationAddress
	[string]$port	
	[string]$communityString
	[string]$userName
	[string]$trapType	 										# trap or Inform
	[string]$engineId											# prefix with 0x - followed by even number of 10 to 64 Hex digits
	[string]$trapSeverities	
	[string]$vcmTrapCategories	
	[string]$enetTrapCategories	
	[string]$fcTrapCategories

}


class eg 
{
	[string]$name
	[string]$logicalInterConnectGroupMapping
	[int]$enclosureCount
	[string]$IPv4AddressingMode
	[string]$IPv4Range
	[string]$IPv6AddressingMode
	[string]$IPv6Range
	[string]$powerMode
	[string]$deploymentMode
	[string]$deploymentNetwork
	[string]$scopes					

}

class le
{
	[string]$name
	[string]$enclosureSerialNumber
	[string]$enclosureName
	[string]$enclosureGroup
	[string]$fwBaseline	
	[Boolean]$fwInstall
	[string]$scopes									

}

class server
{
	[string]$name	
	[string]$serverName	
	[string]$description	
	[string]$model	
	[int]$processorCount	
	[int]$processorCoreCount	
	[string]$processorSpeed	
	[string]$processorType	
	[int]$memory	
	[string]$serialNumber	
	[string]$virtualSerialNumber	
	[string]$profileName	
	[string]$scopes	
}



class spt
{
	[string]$name
	[string]$description
	[string]$serverProfileDescription

	[string]$sht
	[string]$enclosureGroupName
	[string]$affinity

	[Boolean]$manageFirmware
	[string]$firmwareConsistencyChecking
	[string]$firmwareBaselineName
	[string]$firmwareInstallType
	[Boolean]$forceInstallFirmware
	[string]$firmwareActivationType
	[string]$firmwareSchedule
	
	[Boolean]$manageConnections
	[string]$connectionConsistencyChecking

	[string]$localStorageConsistencyChecking 	

	[string]$manageSanStorage						= $False
	[string]$sanStorageConsistencyChecking			
	[string]$hostOSType
	

	[Boolean]$manageBootMode
	[string]$bootModeConsistencyChecking
	[string]$mode	
	[string]$pxeBootPolicy
	[string]$secureBoot


	[Boolean]$manageBootOrder
	[string]$bootOrderConsistencyChecking
	[string]$order
	

	[Boolean]$manageBios
	[string]$biosConsistencyChecking
	[string]$overriddenSettings

	[Boolean]$manageIlo
	[string]$iloConsistencyChecking

	[string]$macType
	[string]$wwnType
	[string]$serialNumberType
	[string]$iscsiInitiatorNameType
	[Boolean]$hideUnusedFlexNics
	[string]$scopes

	
}


class sp
{
	[string]$name
	[string]$description
	[string]$serverProfileTemplate
	[string]$serverHardware
	[string]$empty

	[string]$sht	
	[string]$enclosureGroupName

	[string]$affinity

	[Boolean]$manageFirmware
	[string]$firmwareBaselineName
	[string]$firmwareInstallType
	[Boolean]$forceInstallFirmware
	[string]$firmwareActivationType
	[string]$firmwareSchedule

	[Boolean]$manageLocalStorage
	[Boolean]$manageSanStorage			= $False
	[string]$hostOSType

	[Boolean]$manageBootMode
	[string]$mode	
	[string]$pxeBootPolicy
	[string]$secureBoot

	[Boolean]$manageBootOrder
	[string]$order
	
	[Boolean]$manageBios
	[string]$overriddenSettings

	[Boolean]$manageIlo

	[string]$macType
	[string]$wwnType
	[string]$serialNumberType
	[string]$serialNumber
	[string]$iscsiInitiatorNameType
	[string]$iscsiInitiatorName

	[Boolean]$hideUnusedFlexNics
	[string]$scopes
}

class connection
{
	[string]$profileName
	[string]$name
	[string]$id
	[string]$functionType
	[string]$network
	[string]$portId
	[string]$requestedMbps
	[Boolean]$boot
	[string]$priority
	[string]$bootVolumeSource
	[boolean]$userDefined
	[string]$macType
	[string]$mac
	[string]$wwpnType
	[string]$wwnn
	[string]$wwpn
	[string]$lagName
	[string]$requestedVFs

	
}

class ilosetting
{
	[string]$profileName
	[string]$profileType  						# SPT or SP
	[string]$settingType
	[Boolean]$deleteAdministratorAccount
	[string]$adminPassword					=  '***REDACTED***'

	[string]$userName
	[string]$displayName
	[string]$userPassword					=  '***REDACTED***'
	[string]$userPrivileges

	[string]$directoryAuthentication
	[Boolean]$directoryGenericLDAP
	[string]$iloObjectDistinguishedName
	[string]$directoryPassword				=  '***REDACTED***'
	[string]$directoryServerAddress
	[string]$directoryServerPort
	[string]$directoryUserContext
	[Boolean]$kerberosAuthentication
	[string]$kerberosRealm
	[string]$kerberosKDCServerAddress
	[string]$kerberosKDCServerPort
	[string]$kerberosKeytab
	
	[string]$groupDN
	[string]$groupSID
	[string]$groupPrivileges
	
	[string]$hostName
	
	[string]$primaryServerAddress
	[string]$primaryServerPort
	[string]$secondaryServerAddress
	[string]$secondaryServerPort
	[Boolean]$redundancyRequired
	[string]$groupName
	[string]$certificateName
	[string]$loginName
	[string]$keyManagerpassword				=  '***REDACTED***'
}

class spLocalStorage 
{
	[string]$profileName
	[string]$deviceSlot
	[string]$mode 							# RAID or HBA
	[string]$initialize
	[string]$driveWriteCache
	#[string]$predictiveSpareRebuild
	[string]$logicalDrives
	[string]$id 
	[string]$description
	[string]$raidLevel
	[string]$bootable
	[string]$driveTechnology
	[string]$numPhysicalDrives
	[string]$driveMinSize 
	[string]$driveMaxSize 
	[string]$accelerator
	[string]$eraseData
	[string]$persistent
}

#------------------- Functions
function Build-Header([PSCustomObject] $fromObject)
{
    $PropNames    = [System.Collections.ArrayList]::new()
    
    foreach ( $member in ($fromObject.psobject.members | where membertype -eq 'noteProperty') )
    {
        [void]$PropNames.Add( $member.name)
    }
    

    return $PropNames
}

function Get-NamefromUri([string]$uri, $hostconnection)
{
    $name = ""

    if ($Uri)
    {
        try
        {
            $name   = (Send-OVRequest -uri $Uri  -hostName $hostconnection).name 
        }
        catch
        {
            $name = ""
        }
    }

    return $name
}


function Get-TypefromUri([string]$uri, $hostconnection)
{
    $type = ""

    if ($Uri)
    {
        try
        {
            $type   = (Send-OVRequest -uri $Uri -hostName $hostconnection).Type
        }
        catch
        {
            $type = ""
        }
    }

    return $type
}

Function rebuild-fwISO($BaselineObj)
{

	# ----------------------- Rescontruct FW ISO filename
	# When uploading the FW ISO file into OV, all the '.' chars are replaced with "_"
	# so if the ISO filename is:        SPP_2018.06.20180709_for_HPE_Synergy_Z7550-96524.iso
	# OV will show $fw.ISOfilename ---> SPP_2018_06_20180709_for_HPE_Synergy_Z7550-96524.iso
	# 
	# This helper function will try to re-build the original ISO filename

	$newstr = $null

	switch ($BaselineObj.GetType().Fullname)
	{

		'HPOneView.Appliance.Baseline'
		{

			$arrList = New-Object System.Collections.ArrayList

			$StrArray = $BaselineObj.ResourceId.Split($Underscore)

			ForEach ($string in $StrArray)
			{

				[void]$arrList.Add($string.Replace($dot, $Underscore))

			}
			
			$newstr = "{0}.iso" -f [String]::Join($Underscore, $arrList.ToArray())                

		}

		'HPOneView.Appliance.BaselineHotfix'
		{

			$newStr     = $BaselineObj.FwComponents.Filename

		}

		default
		{

			$newstr = $null

		}

	}

	return $newStr
		
}

Function get-scopes($scopesUri,$hostconnection)
{
	$scopeNames	  				= [System.Collections.ArrayList]::new()
	$scopes 					= ""
	if ($scopesUri)
	{
		$scopedResource 	= send-OVRequest -uri $scopesUri
		$scopeUris 			= $scopedResource.scopeUris
		if ($scopeUris)
		{
			$scopeNames	 	= $scopeUris | % { Get-NamefromUri -uri $_ -hostconnection $hostconnection} 
		}
	}
	if ($scopeNames)
	{
		$scopes			   		= $scopeNames -join $SepChar  
	} 
	return $scopes
}

Function Generate-CustomVarCode ([String]$Prefix, [String]$Suffix, [String]$Value, $indentlevel = 0, $isVar = $True)
{

    if ($isVar)
    {
        $Prefix = '${0}' -f $Prefix
    }

	$Prefix 	= $TAB * $indentlevel + $Prefix
    $VarName    = '{0}{1}' -f $Prefix, $Suffix
    

	if ($Value)
	{
		$len 	= $varName.Length
		$len 	= if ($len -ge $MAXLEN) {$MAXLEN} else { $MAXLEN - $len}
		$pad    = $Space * $len
		$out 	= '{0}{1} = {2}' -f $VarName,$pad, $Value
		
	}
	else 
	{
		$out 	= '{0}' -f $VarName
	} 
	
    return $out
}

function writeToFile ([System.Collections.ArrayList]$code,[string]$file)
{
	if ($code.Count -gt 0)
	{
		[System.IO.File]::WriteAllLines([String]$file, $code, [System.Text.Encoding]::UTF8)
	}
}

# ---------- Internal Helper funtion
function generate-scopeCode( $scopes, $indentlevel = 0)
{
	$arr 		= [system.Collections.ArrayList]::new()
	$arr 		= $scopes.split($SepChar) | % { '"{0}"' -f $_ }
	$scopeList  = '@({0})' -f [string]::Join($Comma, $arr)

	$scopeCode	= ($TAB * $indentLevel) + ('{0}' -f $scopeList)  + ' | % { get-OVScope -name $_ | Add-OVResourceToScope -InputObject $object }'

	[void]$scriptCode.Add($scopeCode)
}

function startBlock($indentlevel = 0, $code = $scriptCode )
{
	[void]$code.Add(($TAB * $indentLevel) + '{')
}


function endBlock($indentlevel = 0)
{
	[void]$scriptCode.Add(($TAB * $indentLevel) + '}')
}

function ifBlock($condition, $indentlevel = 0, $code = $scriptCode )
{

	[void]$code.Add((Generate-CustomVarCode -Prefix  $condition 					-isVar $False -indentlevel $indentlevel) )
	[void]$code.Add(($TAB * $indentLevel) + '{')
}

function endIfBlock ($condition='', $indentlevel = 0, $code = $scriptCode)
{

	[void]$code.Add(($TAB * $indentLevel) + '}' + ' # {0}' -f $condition)
}

function elseBlock($indentlevel = 0, $code = $scriptCode)
{
	[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'else' 						-isVar $False -indentlevel $indentlevel) )
	[void]$code.Add(($TAB * $indentLevel) + '{')
}


function endElseBlock ($indentlevel = 0, $code = $scriptCode)
{
	[void]$code.Add(($TAB * $indentLevel) + '}')
}

function newLine($code = $scriptCode)
{
	[void]$code.Add($CR)
}

function add_to_allScripts ($text , $ps1Files)
{
	[void]$allScriptCode.Add($text)
	[void]$allScriptCode.add($ps1Files)
	[void]$allScriptCode.add($CR)
}


# ----- connect Composer
function connect-Composer([string]$sheetName, [string]$WorkBook, $scriptCode )
{
	$composer 				= get-datafromSheet -sheetName $sheetName -workbook $WorkBook


	$hostName 			= $composer.name
	$userName 			= $composer.Username
	$password 			= $composer.password
	$authDomain 		= if ($NULL -ne $composer.authenticationDomain) {$composer.authenticationDomain} else {'LOCAL'}

	

	newLine
	[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'write-host "`n"'  -isVar $False )) 
	[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Connecting to OneView {0}" ' -f $hostName) -isVar $False ))
	
	ifBlock 	-condition 'if ($global:ConnectedSessions)'
	[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '$d = disconnect-OVMgmt -ApplianceConnection $global:ConnectedSessions' -isVar $false -indentLevel 1))
	endIfBlock 
	generate-credentialCode -username $userName -password $password -component 'OneView' -scriptCode $scriptCode
	[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('connect-OVMgmt -hostname {0} -credential $cred -loginAcknowledge:$True -AuthLoginDomain "{1}" ' -f $hostName,$authDomain) -isVar $False ))
    newLine

}

# -------- strip comment and blank row
function get-datafromSheet([string]$sheetName, [string]$WorkBook)
{
	$List                    = import-Excel -path $WorkBook -workSheetName $sheetName 

	$header                  = Build-Header -fromObject $List[0]
	$firstColumn             = $header[0] 

	$List                    = $List | where { $_.$firstColumn -and ($_.$firstColumn -notmatch '^#')} 
	return $List
}

# -------- generate code for username and password
Function generate-credentialCode ($username, $password, $component,$indentLevel=0, $scriptCode)
{
	[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'userName' -value ( ' "{0}" ' -f $userName ) -indentlevel $indentLevel))
	[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'password' -value ( ' "{0}" ' -f $password ) -indentlevel $indentLevel))

	ifBlock -condition 'if ($password)' -indentlevel $indentLevel
	[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'securePassword' -value '$password | ConvertTo-SecureString -AsPlainText -Force' -indentLevel ($indentLevel + 1) ))
	endIfBlock 	-indentlevel $indentLevel 

	elseBlock 	-indentlevel $indentLevel 
	$value 		= 'Read-Host "{0}: enter password for user {1}" -AsSecureString ' -f $component, $username
	[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'securePassword' -Value $value  -indentlevel ($indentLevel+1)))
	endElseBlock  -indentlevel $indentLevel

	[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'cred ' -Value 'New-Object System.Management.Automation.PSCredential  -ArgumentList $userName, $securePassword'  -indentlevel $indentLevel))
	newLine
	
}

# ---------- firmware Bundle
Function Import-firmwareBaseline([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $scriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  

    foreach ($fw in $List)
    {
		$filename 			= $fw.filename
		$name 				= $fw.name
		if ($filename)
		{
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Adding firmware Baseline {0} "' -f $name) -isVar $False ))
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('Add-OVBaseline -file  "{0}"' -f $filename) -isVar $False ))
            newLine
        }
	}

	if ($List)
	{
		[void]$scriptCode.Add('Disconnect-OVMgmt')
	}

	# ---------- Generate script to file
	writeToFile -code $ScriptCode -file $ps1Files
}


# ---------- Data Center
Function Import-dataCenter([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $scriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  


	foreach ($dc in $List)
	{
		$name 				= $dc.name
		$width				= $dc.width
		$depth 				= $dc.depth
		$millimeters 		= $dc.millimeters
		$defaultVoltage 	= $dc.defaultVoltage
		$powerCosts 		= $dc.powerCosts
		$currency			= $dc.currency
		$coolingCapacity	= $dc.coolingCapacity
		$address1			= $dc.address1
		$address2			= $dc.address2
		$city				= $dc.city
		$state				= $dc.state
		$postCode			= $dc.postCode
		$country			= $dc.country
		$timezone			= $dc.timezone
		$primaryContact		= $dc.primaryContact
		$secondaryContact	= $dc.secondaryContact

		if ($name  -or $width -or $depth )
		{
			newLine
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating datacenter {0} "' -f $name) -isVar $False ))
			$value 					= 'Get-OVDataCenter' + " | where name -eq  '{0}' " -f $name
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'dc' 						-Value $value ))		

			ifBlock -condition 'if ($dc -eq $Null)'
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('# -------------- Attributes for dc {0} ' -f $name) -isVar $False -indentlevel 1))
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'name' 			-Value ("'{0}'" -f $name) -indentlevel 1))

			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'width' 			-Value ('{0}' 	-f $width) -indentlevel 1))
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'depth' 			-Value ('{0}' 	-f $depth) -indentlevel 1))
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'millimeters' 	-Value ('${0}' 	-f $millimeters) -indentlevel 1))

			$nameParam 			= ' -Name "{0}" '		-f $name
			$widthParam 		= ' -Width {0}  ' 		-f $width
			$depthParam 		= ' -Depth {0}  ' 		-f $depth
			$millimetersParam 	= ' -Millimeters:${0}'	-f $millimeters

			$voltageParam  		= $null
			if ($defaultVoltage)
			{
				$voltageParam 	= ' -DefaultVoltage $defaultVoltage'
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'defaultVoltage' 	-Value ('{0}' -f $defaultVoltage) -indentlevel 1))

			}

			$currencyParam  		= $null
			if ($currency)
			{
				$currencyParam 	= ' -Currency $currency'
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'currency' 		-Value ("'{0}'" -f $currency) -indentlevel 1))

			}

			$powerCostsParam  		= $null
			if ($powerCosts)
			{
				$powerCostsParam 	= ' -PowerCosts $powerCosts'
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'powerCosts' 		-Value ('{0}' -f $powerCosts) -indentlevel 1))

			}

			$coolingCapacityParam  = $null
			if ($coolingCapacity)
			{
				$coolingCapacityParam 	= ' -CoolingCapacity $coolingCapacity'
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'coolingCapacity' -Value ('{0}' -f $coolingCapacity) -indentlevel 1))
			}

			$address1Param  = $null
			if ($address1)
			{
				$address1Param 	= ' -Address1 $address1'
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'address1' 		-Value ("'{0}'" -f $address1) -indentlevel 1))
			}

			$address2Param  = $null
			if ($address2)
			{
				$address2Param 	= ' -Address2 $address2'
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'address2' 		-Value ("'{0}'" -f $address2) -indentlevel 1))
			}

			$cityParam  = $null
			if ($city)
			{
				$cityParam 	= ' -City $city'
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'city' 			-Value ("'{0}'" -f $city) -indentlevel 1))
			}

			$stateParam  = $null
			if ($state)
			{
				$stateParam 	= ' -State $state'
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'state' 			-Value ("'{0}'" -f $state) -indentlevel 1))
			}

			$postCodeParam  = $null
			if ($postCode)
			{
				$postCodeParam 	= ' -PostCode $postCode'
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'postCode' 		-Value ("'{0}'" -f $postCode) -indentlevel 1))
			}

			$countryParam  = $null
			if ($country)
			{
				$postCodeParam 	= ' -Country $country'
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'country' 		-Value ("'{0}'" -f $country) -indentlevel 1))
			}

			$timezoneParam  = $null
			if ($timezone)
			{
				$timezoneParam 	= ' -TimeZone $timezone'
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'timezone' 		-Value ("'{0}'" -f $timezone) -indentlevel 1))
			}

			$primaryContactParam  = $null
			if ($primaryContact)
			{
				$primaryContactParam 	= ' -PrimaryContact $primaryContact'
				$value					= 'Get-OVRemoteSupportContact -Name "{0}" ' -f $primaryContact
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'primaryContact' 	-Value  $value -indentlevel 1))
			}

			$secondaryContactParam  = $null
			if ($secondaryContact)
			{
				$secondaryContactParam 	= ' -SecondaryContact $secondaryContact'
				$value					= 'Get-OVRemoteSupportContact -Name "{0}" ' -f $secondaryContact
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'secondaryContact' 	-Value $value  -indentlevel 1))
			}
			

			# Code
			# ensure that there is no space after backstick(`)
			newLine

			$prefix		= 'New-OVDatacenter {0}{1}{2}{3} `' -f $nameParam,$widthParam,$depthParam,$millimetersParam
			[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix $prefix -isVar $False -indentLevel 1)) 

			$prefix 	= '{0}{1}{2}{3} `' 	-f $voltageParam, $powerCostsParam, $currencyParam, $coolingCapacityParam 
			[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix $prefix -isVar $False -indentLevel 2))

			$prefix 	= '{0}{1}{2}{3}{4}{5}{6} `' 	-f $address1Param, $address2Param, $cityParam, $stateParam, $postCodeParam,  $countryParam , $timezoneParam
			[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix $prefix -isVar $False -indentLevel 2))

			$prefix 	= '{0}{1} `' 	-f $primaryContactParam, $secondaryContactParam
			[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix $prefix -isVar $False -indentLevel 2))

			newLine # to end the command
			endifBlock -condition 'if ($dc -eq $Null)'

			# Skip creating because resource already exists
			elseBlock
			[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $name + ' already exists.') -isVar $False -indentlevel 1 ))
			endElseBlock


		newLine
		}

	}

	if ($List)
	{
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	}

	# ---------- Generate script to file
	writeToFile -code $ScriptCode -file $ps1Files
}


# ---------- Rack
Function Import-rack([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $scriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  


	foreach ($rack in $List)
	{
		$dcName 			= $rack.name
		$rackSN 			= $rack.rackSerialNumber
		$x 					= $rack.xCoordinate
		$y 					= $rack.yCoordinate
		$millimeters		= $rack.millimeters

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Adding rack {0} to datacenter {1} "' -f $rackSN, $dcName) -isVar $False ))
		newLine

		$value 				= 'Get-OVDataCenter | where name -match "{0}" ' 	-f $dcName
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'dc' 						-Value $value ))	
		$value 				= 'Get-OVRack | where serialNumber -match "{0}" ' -f $rackSN
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'rack' 					-Value $value ))
		
		newLine

		ifBlock 	-condition 'if ( ($dc -ne $Null) -and ($rack -ne $Null) )' 
		$value 				= '$rack.Uri'
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'rackUri' -Value $value -indentlevel 1 ))

		#$value 				=  '$dc.contents | where resourceUri -match $rackUri'
		#[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'rack_in_dc' -Value $value -indentlevel 1 ))

		ifBlock 	-condition 'if ($null -eq ($dc.contents | where resourceUri -match $rackUri) )'  -indentlevel 1
		# Code here
		$dcParam  			= ' -DataCenter $dc ' 
		$inputParam 		= ' -InputObject $rack '
		$coordParam 		= ' -X {0} -Y {1} -Millimeters:${2}'		-f $x, $y, $millimeters
		$prefix 			= 'Add-OVRackToDataCenter {0}{1}{2} ' 	-f $inputParam,$dcParam, $coordParam
		
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix $prefix -isVar $False -indentLevel 2)) 
		endIfBlock 	-indentlevel 1

		# Rack already defined in dc
		elseBlock 	-indentlevel 1
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $rackSN + ' already defined in data center.') -isVar $False -indentlevel 2 ))
		endElseBlock -indentlevel 1

		endIfBlock 	-condition '$dc -ne $Null and $rack $ne $Null' 

		# Data center not existed
		elseBlock
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW ' + "datacenter '{0}'" -f $dcName + " or rack '{0}'" -f $rackSN  + ' do not exist. Define datacenter first') -isVar $False -indentlevel 1 ))
		endElseBlock



		# Code
		# ensure that there is no space after backstick(`)
		newLine



	}

	if ($List)
	{
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	}

	# ---------- Generate script to file
	writeToFile -code $ScriptCode -file $ps1Files
}


# ---------- proxy
Function Import-proxy([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $scriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  


	foreach ($proxy in $List)
	{
		$name 				= $proxy.server
		$protocol 			= $proxy.protocol
		$port 				= $proxy.port
		$username 			= $proxy.Username

		
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Importing proxy "' -f $name) -isVar $False ))

		$hostNameParam 		= ' -Hostname "{0}"' 	-f $name
		$https 				= if ($protocol -eq 'https') {'$True'} else {'$False'}
		$httpsParam 		= ' -Https:{0}' 		-f $https
		$portParam 			= ' -Port {0}'			-f $port

		$userParam 			= $null
		if ($username)
		{
			$value 			= 'Read-Host "Proxy Setting: enter password for user {0}" -AsSecureString ' -f $username
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'securepass' -value  $value))
			$userParam 		= ' -Username "{0}" -password $securepass ' -f $username
		}
		# Code
		# ensure that there is no space after backstick(`)
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('Set-OVApplianceProxy {0}{1}{2}{3}' -f $hostNameParam,$portParam,$httpsParam,$userParam) -isVar $False ))
		newLine

	}

	if ($List)
	{
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	}
	# ---------- Generate script to file
	writeToFile -code $ScriptCode -file $ps1Files
}


# ---------- proxy
Function Import-TimeLocale([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $scriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  


	foreach ($time in $List)
	{
		$locale 			= $time.locale
		$ntpServers 		= $time.ntpServers
		$pollingInterval 	= $time.pollingInterval
		$syncWithHost		= $time.syncWithHost

		$ntpServers 		= "@('" + $ntpServers.replace($SepChar, "'$Comma'") +  "')" 
		if ($syncWithHost -like 'True')
		{
			$syncParam 		= ' -syncwithHost'
		}

		if ($pollingInterval)
		{
			$pollParam 		= " -PollingInterval {0}" -f $pollingInterval
		}

		# Code
		# ensure that there is no space after backstick(`)
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('Set-OVApplianceDateTime -locale "{0}" -ntpServers {1} {2} {3}' -f $locale,$ntpServers,$syncParam,$pollParam) -isVar $False ))
		newLine

	}

	if ($List)
	{
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	}

	# ---------- Generate script to file
	writeToFile -code $ScriptCode -file $ps1Files
}


 Function Import-backup([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
 {
	 $scriptCode            = [System.Collections.ArrayList]::new()
	 $cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	 connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	 
	 $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	 if ($List)
	 {
		 [void]$scriptCode.Add($codeComposer)
	 }
	 
	   

	 foreach ($bkp in $List)
	 {
		$enabled 			  = $bkp.enabled
		$server 			  = $bkp.remoteServerName
		$dir 				  = $bkp.remoteServerDir
		$protocol 		      = $bkp.protocol
		$username 			  = $bkp.Username
		$password 			  = $bkp.Password
		$interval			  = $bkp.scheduleInterval
		$days 				  = $bkp.scheduleDays
		$time  				  = [string]$bkp.scheduleTime
		$publicKey			  = $bkp.remoteServerPublicKey

		if ($enabled -like 'True')
		{

			if ($server)
			{
				$protocol 	      = if ($protocol -eq $Null ) {'SCP'} else {$protocol}
				$remoteParam      = ' -Hostname "{0}"  -protocol "{1}" -HostSSHKey $hostSSHKey' -f $server, $protocol, $publicKey
				if ($dir)
				{
					$remoteParam  += ' -Directory "{0}" ' 	-f $dir
				}
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'hostSSHKey' -value ("'{0}'" -f $publicKey) ))

				if ($NULL -eq $password)
				{
					$value 			= 'Read-Host "Backup Config Setting --> enter password for user {0}" -AsSecureString ' -f $username
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'securepass' -value  $value))
				}
				else
				{
					$value 			= "'$password' | ConvertTo-SecureString -AsPlainText -Force "
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'securepass' -value  $value))
				}

				$userParam 		= ' -Username "{0}" -password $securepass ' -f $username

				$scheduleParam  = ' -Interval "{0}" -Time "{1}" ' -f $interval,$time
				if ( ($interval -match 'Weekly') -and ($Days) )
				{
					$Days 		= "@('" + $Days.replace($SepChar, "'$Comma'") +  "')" 
					$scheduleParam  += ' -Days {0}' -f $Days
				}

			}


			# Code
			# ensure that there is no space after backstick(`)
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('Set-OVAutomaticBackupConfig {0} `' -f $remoteParam) -isVar $False ))
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0} {1}' -f $userParam, $scheduleParam) -isVar $False -indentlevel 1))
			newLine
		}
		else # Disable backup
		{
			# Code
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'Set-OVAutomaticBackupConfig  -Disabled:$True '  -isVar $False ))
		}

	 }
	 
	 if ($List)
	 {
		 [void]$ScriptCode.Add('Disconnect-OVMgmt')
	 }
  # ---------- Generate script to file
  writeToFile -code $ScriptCode -file $ps1Files
 }


 Function Import-repository([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
 {
	 $scriptCode             = [System.Collections.ArrayList]::new()
	 $cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	 connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	 
	 $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	 if ($List)
	 {
		 [void]$scriptCode.Add($codeComposer)
	 }
	 
   
	 foreach ($repo in $List)
	 {
		$name 				  	= $repo.name
		$hostName 		  		= $repo.hostName
		$directory				= $repo.directory
		$protocol 				= $repo.protocol
		$username 				= $repo.Username
		$password				= $repo.password
		$certificate 			= $repo.certificate


		$hostParam 				= " -HostName  '{0}' " -f $hostName
		$httpParam 				= if ($protocol -like 'True') 	{' -http'} 									else {''}
		$dirParam 				= if ($directory) 				{ ' -Directory "{0}" ' -f $directory}  		else {''}
		$certParam 				= if ($certificate)				{ ' -Certificate "{0}" ' -f $certificate}	else {''}
		

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating external repository {0} "' -f $name) -isVar $False ))

		if ($username)
		{
			generate-credentialCode -username $username -password $password -component 'REPOSITORY' -scriptCode $scriptCode
			$credParam 		= ' -Credential $cred'
	 	}

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('New-OVExternalRepository -Name "{0}" {1}{2}{3}{4}{5}' -f $name, $hostParam, $httpParam, $dirParam, $credParam, $certParam) -isVar $False ))
	

		# Code
	 }
	 
	 if ($List)
	 {
		 [void]$ScriptCode.Add('Disconnect-OVMgmt')
	 }
	# ---------- Generate script to file
	writeToFile -code $ScriptCode -file $ps1Files
 }

 
 Function Import-scope([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
 {
	 $scriptCode             = [System.Collections.ArrayList]::new()
	 $cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	 connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	 
	 $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	 if ($List)
	 {
		 [void]$scriptCode.Add($codeComposer)
	 }
	 
   
	 foreach ($scope in $List) 
	 {
		$name 				= $scope.name
		$description		= $scope.description
		$resource 			= $scope.resource



		$descParam 			= if ($description) { ' -Description "{0}" ' -f $description} else {''}
		
		newLine
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating scopes {0} "' -f $name) -isVar $False ))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix username -Value ("'{0}'" -f $name)  ))

		ifBlock 		-condition ('if ($null -eq (get-OVScope | where name -eq "{0}" ))' -f $name)
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('New-OVScope -name "{0}" {1} ' -f $name, $descParam) -isVar $False -indentlevel 1))
		
		newLine
		
		if ($resource)
		{
			$resArray  		= $resource.Split($SepChar)
			$index 			= 1
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix resources -value '@()' -indentLevel 1) )
			foreach ($res in $ResArray)
			{
				$resType, $resName 	= $res.Split(';')
                $resType 			= $resType.replace('type=', '').Trim()
                if ($resName)
				{
                    $resName 			= '"' + $resName.replace('name=', '').Trim() + '"'  # etract name and surround with quotes
                }

				$value 	= 'Get-OV{0} | where name -eq {1}' -f $resType, $resName
				$prefix = "res$Index"
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix $prefix -value $value -indentlevel 1 ))

				ifBlock		-condition ('if (${0})' -f $prefix) -indentlevel 1
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix resources -value ('$resources + ${0}' -f $prefix) -indentlevel 2 ))
				endIfBlock -indentlevel 1 

				ifBlock 	-condition 'if ($resources)'
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('get-OVScope -name "{0}" | Add-OVResourceToScope -InputObject $resources ' -f $name) -isVar $False -indentlevel 2))
				endIfBlock -indentlevel 1
				newLine
			}


		}
		
		
		endIfBlock		-condition '$null -eq (get-OVScope....'
		# Skip creating because resource already exists

		elseBlock
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $name + ' already exists.') -isVar $False -indentlevel 1 ))
		endElseBlock


	 }
	 
	if ($List)
	{
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	}
   # ---------- Generate script to file
   writeToFile -code $ScriptCode -file $ps1Files
 }


 Function Import-snmpTrap([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
 {
	 $scriptCode             = [System.Collections.ArrayList]::new()
	 $cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	 
	 
	 $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	 if ($List)
	 {
		$List 				= $List | where source -eq 'Appliance' # Select Appliance snmp user only
		connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
		create-snmpTrap -list $List -isSnmpAppliance $True
		newLine
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	 }
   # ---------- Generate script to file
   writeToFile -code $ScriptCode -file $ps1Files
 }		 

 Function create-snmpTrap($List, $isSnmpAppliance=$True, $indent=0 )
 {
	$TextInfo 			= (Get-Culture).TextInfo
	$snmpTraps			= [System.Collections.ArrayList]::new()
	$index 				= 0 
	$isSnmpLig 			= -not $isSnmpAppliance
	 
 
	 foreach ($trap in $List) 
	 {
		$Index++

		$format 					= $trap.format
		$destinationAddress			= $trap.destinationAddress
		$port 						= $trap.port              
		$communityString 			= $trap.communityString
		$snmpV3User 				= $trap.userName

		$formatParam = $destinationParam = $portParam = $trapTypeParam = $engineIdParam = $communityParam  = $snmpV3userParam = ""
		$destinationParam 	= if ($destinationAddress) 	{ ' -Destination "{0}" ' 	-f $destinationAddress 	} 		else {""}
		$portParam 			= if ($port)				{ ' -Port {0} ' 		 	-f $port				} 		else {""}
		$communityParam		= if ($communityString)		{ ' -Community {0} '		-f $communityString		} 		else {""}  

		if ($isSnmpLig -and ($format -eq 'snmpV3') )
		{

			$trapType 				= $TextInfo.ToTitleCase($trap.trapType.tolower())
			$engineId 				= $trap.engineId -replace '0x' , '10x'
			$trapTypeParam 			= if ($trapType)	{ ' -NotificationType "{0}" '	-f $trapType			} 		else {""}
			$engineIdParam			= if ($engineId)	{ ' -EngineID "{0}" ' 	 	-f $engineId			} 		else {""} 
		}
		
		$formatParam 				= if ($format -and $isSnmpLig) { " -SnmpFormat $format " }						else { " -Type $format " }

		

		if ($snmpV3User -and ($format -eq 'snmpV3') )
		{
			if ($isSnmpAppliance)
			{
				# Use Get-OVsnmpV3user to get the user object
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'snmpV3user' 	-value ('Get-OVSnmpV3User | where UserName -eq "{0}" ' -f $snmpV3User) -indentlevel $indent))	
			}
			else 
			{
				# Use snmpV3Users / snmpV3UserNames variable in create_snmpV3user
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'this_index' -Value ('[array]::IndexOf($snmpV3Usernames, "{0}" )' -f $snmpV3User) -indentlevel $indent))	
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'snmpV3user' -Value '$snmpV3Users[$this_index]' -indentlevel $indent								  ))
			}

			$snmpV3userParam = ' -SnmpV3user $snmpV3user '
		}

		$new_trap		= '$trap{0}' -f $Index	
		$trapCmd 		= if ($isSnmpAppliance) { 'New-OVApplianceTrapDestination '} else {'New-OVSnmpTrapDestination '}

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ("write-host -foreground CYAN '----- Creating snmp trap {0} '"  -f $new_trap)  -isVar $False -indentlevel $indent))
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('{0}' -f $new_trap) -value ('{0}{1}{2}{3}{4}{5}{6}{7}' -f $trapCmd, $formatParam,$destinationParam,$portParam,$trapTypeParam,$engineIdParam,$communityParam,$snmpV3userParam ) -isVar $false -indentlevel  $indent ))
		newLine
		$snmpTraps		+= $new_trap		

	 }

	 if ($isSnmpLig)
	 {
		 newLine
		 [void]$scriptCode.Add(( Generate-CustomVarCode -Prefix 'snmpTraps' -value ('@({0})' -f ($snmpTraps  -join $COMMA) ) -indentlevel  $indent ))
	 }	
	 

   # ---------- Generate script to file
   writeToFile -code $ScriptCode -file $ps1Files
 }	


 function create-snmpV3User($List, $isSnmpAppliance=$True, $indent=0)
 {

	 $snmpv3Users			= [System.Collections.ArrayList]::new()
	 $snmpv3UserNames		= [System.Collections.ArrayList]::new()
	 $index 				= 0 
	 $isSnmpLig 			= -not $isSnmpAppliance

	 if ($List -and $isSnmpLig)
	 {
		newLine
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'snmpv3Users' 	-value	'[System.Collections.ArrayList]::new()' -indentlevel $indent ) )
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'snmpv3Usernames' -value	'[System.Collections.ArrayList]::new()' -indentlevel $indent ) )
	 }
     foreach ($snmp in $List) 
	 {
		$snmpV3User					= $snmp.userName
		$securityLevel 				= $snmp.securityLevel
		$authProtocol 				= $snmp.authProtocol
		$authPassword 				= $snmp.authPassword
		$privProtocol 				= $snmp.privacyProtocol
        $privPassword 				= $snmp.privacyPassword
		
		
        # ----- snmpV3 user
		if ($snmpV3User)
		{
			$index++ # Next user

            [void]$scriptCode.Add($CR)
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating snmpV3 user {0} " ' -f $snmpV3User)  -isVar $False -indentlevel $indent))

			if ($isSnmpAppliance)
            {
				$condition              = 'if ($null -eq (Get-OVSnmpV3User | where name -eq "{0}"))' -f $snmpV3User
            	[void]$scriptCode.Add((Generate-CustomVarCode -Prefix  $condition 	-isVar $False -indentlevel $indent  ) )
            	[void]$scriptCode.Add('{')

				$applianceSnmpParam 	= ' -ApplianceSnmpUser '
				$indent 				+= 1
			}
			$secParam 				= $null
			switch ($securityLevel)
			{
				'AuthOnly'		
					{

						[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'authPassword' -value ('"{0}" | ConvertTo-SecureString -AsPlainText -Force' -f $authPassword) -indentLevel $indent))
						$secParam	= ' -SecurityLevel "{0}" -AuthProtocol "{1}" -AuthPassword $authPassword' -f $securityLevel, $authProtocol
					}
				'AuthAndPriv'
					{
						[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'authPassword' -value ('"{0}" | ConvertTo-SecureString -AsPlainText -Force' -f $authPassword) -indentLevel  $indent ))
						[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'privPassword' -value ('"{0}" | ConvertTo-SecureString -AsPlainText -Force' -f $privPassword)  -indentLevel  $indent))
						$secParam	= ' -SecurityLevel "{0}" -AuthProtocol "{1}" -AuthPassword $authPassword -PrivProtocol "{2}" -PrivPassword $privPassword' -f $securityLevel, $authProtocol, $privProtocol 
					}										
			}

			$new_snmpv3User 		= '$snmpv3User{0}' -f $Index
			[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('{0}' -f $new_snmpv3user) -value ('new-OVSnmpV3User {0} -UserName "{1}" {2}' -f $applianceSnmpParam, $snmpV3User, $secParam ) -isVar $false -indentlevel  $indent ))
			$snmpv3Users 		+= $new_snmpv3User
			$snmpv3Usernames	+= $snmpv3User			# Collect user names

			if ($isSnmpAppliance)
            {
				[void]$scriptCode.Add('}')
				# Skip creating because resource already exists
            	[void]$scriptCode.Add('else')
            	[void]$scriptCode.Add('{')
				[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $snmpV3User + ' already exists.') -isVar $False -indentlevel  $indent ))
				[void]$scriptCode.Add('}')
			}


			
		}
	 
    
	 }
	 
	 if ($snmpv3Users -and $isSnmpLig)
	 {
		 newLine
		 [void]$scriptCode.Add(( Generate-CustomVarCode -Prefix 'snmpv3Users' -value ('@({0})' -f ($snmpv3Users  -join $COMMA) ) -indentlevel  $indent ))
		 
		 # Set collection of snmp v3 user names
		 $snmpv3UserNames 	= $snmpv3UserNames | % { "'{0}'" -f  $_ }  	# Add prefix $
		 [void]$scriptCode.Add(( Generate-CustomVarCode -Prefix 'snmpv3Usernames' -value ("@({0})" -f ($snmpv3UserNames  -join $COMMA) ) -indentlevel  $indent ))
	 }

	 

 }


 Function Import-snmpV3User([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
 {
	 $scriptCode            = [System.Collections.ArrayList]::new()
	 $cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	 connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	 
	 $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	 if ($List)
	 {
		 $List 				= $List | where source -eq 'Appliance' # Select Appliance snmp user only
		 [void]$scriptCode.Add($codeComposer)
		 create-snmpV3User -list $List -isSnmpAppliance $True
		 newLine
		 [void]$ScriptCode.Add('Disconnect-OVMgmt')
	 }
 
   # ---------- Generate script to file
   writeToFile -code $ScriptCode -file $ps1Files
 }

 Function Import-remoteSupport([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
 {
	 $scriptCode             = [System.Collections.ArrayList]::new()
	 $cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	 connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	 
	 $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	 if ($List)
	 {
		 [void]$scriptCode.Add($codeComposer)
		 $enable 			= $List.enabled
		 $companyName 		= $List.companyName
		 $username 			= $List.insightOnlineUsername
		 $password 			= $List.insightOnlinePassword
		 $optimizeOptIn 	= $List.optimizeOptIn
	 
		 if ($enable -eq 'True')
		 {
			$enableParam 		=  ' -enable ' 
			$optParam 			= if ($optimizeOptIn -eq 'True')	{ ' -OptimizeOptIn ${0}' -f $optimizeOptIn} else {''}
				

			generate-credentialCode -username $username -password $password -scriptCode $scriptCode
			$credParam 			= ' -InsightOnlineUsername "{0}" -InsightOnlinePassword $securepassword ' -f $username
			[void]$ScriptCode.Add( (Generate-CustomVarCode -prefix ('Set-OVRemoteSupport -enable -CompanyName "{0}" {1} {2}' -f $companyName, $credParam, $optParam ) ))
		 }
		 else 
		 {
			[void]$ScriptCode.Add( (Generate-CustomVarCode -prefix 'Set-OVRemoteSupport -disable ' ))
		 }

	   [void]$ScriptCode.Add('Disconnect-OVMgmt')
	 }
  # ---------- Generate script to file
  writeToFile -code $ScriptCode -file $ps1Files

 }


 ####### TO BE COMPLETED
 Function Import-ligsnmp([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
 {
	 $scriptCode             = [System.Collections.ArrayList]::new()
	 $cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	 connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	 
	 $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	 if ($List)
	 {
		 [void]$scriptCode.Add($codeComposer)
	 }
	 
   
	 foreach ($snmp in $List) 
	 {
		$communityString 			= $snmp.communityString
		$trapDestination			= $snmp.trapDestination # []
		$port 						= $snmp.port            
		$snmpFormat 				= $snmp.snmpFormat 
		$trapSeverities				= $snmp.trapSeverities  
		$enetTrapCategories			= $snmp.enetTrapCategories 
		$fcTrapCategories			= $snmp.fcTrapCategories  
		$notificationType			= $snmp.notificationType 
		$engineId 					= $snmp.engineId  
		$trapsnmpV3User 			= $snmp.trapsnmpV3User


		newLine
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'write-host -foreground CYAN "----- Configuring OV snmp {0} "' -isVar $False ))
		newLine

		# ------------ snmp Read Community string
		if ($communityString)
		{
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'write-host -foreground CYAN "----- Importing snmp read Community string "'  -isVar $False ))
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('Set-OVSnmpReadCommunity -name "{0} "' -f $readCommunity)  -isVar $False ))
			newLine
		}

		# ------------ snmp trap

		$destArray 				= if ($trapDestination) 	{ $trapDestination.Split($SepChar) } 	else {$null}
		$portArray 				= if ($port) 				{ $port.Split($SepChar) } 				else {$null}
		$fmtArray 				= if ($snmpFormat) 			{ $snmpFormat.Split($SepChar) } 		else {$null}
		$sevArray 				= if ($trapSeverities) 		{ $trapSeverities.Split($SepChar) } 	else {$null}
		$enetArray 				= if ($enetTrapCategories) 	{ $enetTrapCategories.Split($SepChar) } else {$null}
		$fcArray 				= if ($fcTrapCategories) 	{ $fcTrapCategories.Split($SepChar) } 	else {$null}
		$notifArray 			= if ($notificationType) 	{ $notificationType.Split($SepChar) } 	else {$null}
		$v3UserArray 			= if ($trapsnmpV3User) 		{ $trapsnmpV3User.Split($SepChar) } 	else {$null}


		if ($trapDestination)
		{
			
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'write-host -foreground CYAN "----- Importing snmp trap destination"'  -isVar $False ))
			for ($i=0; $i -lt $destArray.Count; $i++)
			{
				$destinationParam 	= ' '
			}

			newLine
		}





		endBlock
		# Skip creating because resource already exists
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix 'else'-isVar $False))
		startBlock
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $name + ' already exists.') -isVar $False -indentlevel 1 ))
		endBlock


	 }
	 
	if ($List)
	{
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	}
   # ---------- Generate script to file
   writeToFile -code $ScriptCode -file $ps1Files
 }		 

 ####### TO BE COMPLETED


 Function Import-user([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
 {
	 $scriptCode             = [System.Collections.ArrayList]::new()
	 $cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	 connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	 
	 $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	 if ($List)
	 {
		 [void]$scriptCode.Add($codeComposer)
	 }
	 
	 $ovList				= $List | where type -eq 'OV'
	 $rsList 				= $List | where type -eq 'RS'



	 # -------------------- Remote support contacts
	 foreach ($rs in $rsList)
	 {
		 $firstName 		= $rs.name
		 $lastName			= $rs.fullName
		 $email				= $rs.emailAddress
		 $primary 			= $rs.officePhone
		 $default 			= $rs.default
		 $language 			= $rs.language
		 $notes				= $rs.notes

		 $defaultParam 		= if ($default -eq 'True') 	{ ' -Default'} 						else {''}
		 $languageParam 	= if ($language ) 			{ ' -Language "{0}"' -f $language}	else {''}
		 $notesParam 		= if ($notes ) 				{ ' -Notes  "{0}"' -f $notes} 		else {''}

		 $nameParam 		= ' -Firstname "{0}" -Lastname "{1}" -email "{2}" -primary {3} ' -f $firstName, $lastName, $email, $primary

		 newLine
		 [void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating Remote Support contacts {0} "' -f $firstName) -isVar $False ))

		 ifBlock 		-condition ('if ($null -eq (Get-OVRemoteSupportContact | where email -eq "{0}" ))' -f $email)
		 [void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('new-OVRemoteSupportContact {0}{1}{2}{3}' -f $defaultParam, $nameParam,$languageParam, $notesParam) -isVar $False -indentlevle 1))		 
		 endIfBlock

		 #Skip creating because resource already exists
		 elseBlock
		 [void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $email + ' already exists.') -isVar $False -indentlevel 1 ))
		 endElseBlock

	 }
	 
	 # -------------------- OV users

	 foreach ($user in $ovList)
	 {
		$name 				  	= $user.name
		$password				= $user.password
		$fullName 		  		= $user.fullName
		$emailAddress			= $user.emailAddress
		$officePhone 			= $user.officePhone
		$mobilePhone 			= $user.mobilePhone
		$roles 	 				= $user.roles
		$permissions 		 	= $user.permissions

		newLine
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating local users {0} "' -f $name) -isVar $False ))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix username -Value ("'{0}'" -f $name)  ))

		ifBlock			-condition ('if ($null -eq (get-OVUser | where userName -eq "{0}" ))' -f $name) 
		generate-credentialCode -component 'Users' -username $name -password $password -scriptCode $scriptCode  -indentlevel 1

		$fullNameParam 			= if ($fullName)		{ ' -FullName "{0}" ' 		-f $fullName }			else {''}
		$emailParam 			= if ($emailAddress)	{ ' -EmailAddress "{0}" ' 	-f $emailAddress }		else {''}
		$officeParam 			= if ($officePhone) 	{ ' -OfficePhone "{0}" ' 	-f $officePhone }		else {''}
		$mobileParam 			= if ($mobilePhone) 	{ ' -MobilePhone "{0}" ' 	-f $mobilePhone }		else {''}

		$rolesParam 			= $null
		if ($roles)
		{
			$roles 					= "@('" + $roles.Replace($SepChar,"'$comma'") + "')"
			$rolesParam 			= ' -Roles $roles'
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'roles' 	-Value $roles -indentlevel 1))
		}


		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix scopePermissions -Value '@()' -indentlevel 1  ))
		if ($permissions)
		{
			# Extract the scope name to build a query
			$permArray		= $permissions.Split($Sepchar)
			$index 			= 1 
			foreach ($el in $permArray)
			{
				$scopeName 	= $el.Split($Equal)[-1].Trim()
				$role 		= $el.Split(';')[0].Trim()
				$role 		= $role.replace($Equal, "$Equal'") + "'"    # Add quote around role name

				$scopeIndex = "scope$Index"
				$value 		= 'Get-OVScope | where name -eq "{0}" ' -f $scopeName 
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix $scopeIndex -value $value -indentlevel 1))

				ifBlock 	-condition ('if ($null -ne ${0})' -f $scopeIndex) 	-indentlevel 1
				$spIndex 	= "sp$Index"
				$value 		= '@{' + '{0};scope=${1}' -f $role,$scopeIndex + '}'
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix $spIndex -value $value -indentlevel 2))
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix scopePermissions -Value ('$scopePermissions + ${0}' -f $spIndex) -indentlevel 2  ))
				endIfBlock -indentlevel 1

				$Index++
			}

			
		}

		# Ensure there is no space after backtick (`)

		ifBlock			-condition 'if ($scopePermissions)' -indentlevel 1
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('New-OVUser -username "{0}" -password "{1}" `' -f $name, $password) -isVar $False -indentlevel 2))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0}{1}{2}{3}`' -f $fullNameParam, $emailParam ,$officeParam, $mobileParam) -isVar $False -indentlevel 4))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0} -ScopePermissions $scopePermissions' -f $rolesParam) -isVar $False -indentlevel 4))
		endIfBlock		-indentlevel 1

		elseBlock 		-indentlevel 1
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('New-OVUser -username "{0}" -password "{1}" `' -f $name, $password) -isVar $False -indentlevel 2))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0}{1}{2}{3}`' -f $fullNameParam, $emailParam ,$officeParam, $mobileParam) -isVar $False -indentlevel 4))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0}' -f $rolesParam) -isVar $False -indentlevel 4))		
		endElseBlock 	-indentlevel 1

		endIfBlock		-condition 'Get-OVUser...'

		# Skip creating because resource already exists
		elseBlock
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $name + ' already exists.') -isVar $False -indentlevel 1 ))
		endElseBlock

	 }
	 
	if ($List)
	{
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	}
   # ---------- Generate script to file
   writeToFile -code $ScriptCode -file $ps1Files
 }


 Function Import-addressPool([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
 {
	 $scriptCode            = [System.Collections.ArrayList]::new()
	 $cSheet, $sheetName   	= $sheetName.Split($SepChar)       # composer

	 connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	 
	 $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	 if ($List)
	 {
		 [void]$scriptCode.Add($codeComposer)
	 }


	 foreach ($addr in $List)
	 {
		$name 				  	= $addr.name
		$poolType			  	= $addr.poolType
		$rangeType				= $addr.rangeCategory
		$deleteGenerated		= [Boolean]($addr.deleteGenerated)
		$startAddress 			= $addr.startAddress
		$endAddress 			= $addr.endAddress
		$networkId	 			= $addr.networkId
		$subnetmask 			= $addr.subnetmask
		$gateway 				= $addr.gateway
		$dnsServers 			= $addr.dnsServers
		$domain 				= $addr.domain

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating address pools {0} "' -f $name) -isVar $False ))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('# -------------- Attributes for address Pools "{0}"' -f $name) -isVar $False -indentlevel 1))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'proceed' 		-Value '$False' ) )
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'poolType' 		-Value ('"{0}"' 	-f $poolType) ))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'rangeType' 		-Value ('"{0}"' 	-f $rangeType) ))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'startAddress' 	-Value ('"{0}"' 	-f $startAddress) ))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'endAddress' 		-Value ('"{0}"' 	-f $endAddress) ))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'deleteGenerated' -Value ('${0}' 		-f $deleteGenerated) ))

		if ($poolType -like 'ip*')
		{
			
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'name' 		-Value ('"{0}"' 	-f $name) ))
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'networkId' 	-Value ('"{0}"' 	-f $networkId) ))
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'subnetmask' 	-Value ('"{0}"' 	-f $subnetmask) ))
			$gwParam 		= $null
			if ($gateway)
			{
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'gateway' 	-Value ('"{0}"' 	-f $gateway) ))
				$gwParam 	= '-Gateway $gateway '
			}
			$dnsParam 		= $null
			if ($dnsServers)
			{
				$value 		= "@('" + $dnsServers.replace($SepChar, "'$Comma'") + "')" 
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'dnsServers' 	-Value $value ))
				$dnsParam 	= ' -DNSServers $dnsServers '
			}
			$domainParam 	= $null
			if ($domain)
			{
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'domain' 	-Value ('"{0}"' 	-f $domain) ))
				$domainParam = ' -Domain $domain'
			}

		}

		if ($poolType -like "ip*")
		{
			ifBlock		-condition 'if ($poolType -like "ip*")' 
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'subnet' -Value (" Get-OVAddressPoolSubnet | where networkId -eq '{0}'" -f $networkId ) -indentlevel 1) )

			ifBlock		-condition 'if ($subnet -ne $null) ' -indentlevel 1
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'addressPool' -Value 'Get-OVAddressPoolRange | where subnetUri -match ($subnet.uri)' -indentlevel 2) )
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'proceed' -value ('$null -eq ($addressPool.startStopFragments.startAddress -ne $startAddress)')  -indentlevel 2 ) )
			endIfBlock  -condition '$subnet....'	-indentlevel 1
			
			elseBlock		-indentlevel 1 	# generate new subnet
			$value 			= 'new-OVAddressPoolSubnet -NetworkId $networkId -SubnetMask $subnetMask ' + $gwParam + $dnsParam + $domainParam
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'subnet' -value  $value -indentlevel 2 ) )
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'proceed' -Value '$True' -indentlevel 2 ) )
			endElseBlock	-indentlevel 1 
			endIfBlock	-condition '$poolType....' 
		}
		else 
		{
			$value = 'Get-OVAddressPoolRange| where {($_.name -eq $poolType) -and ($_.rangeCategory -eq $rangeType)} '
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'addressPool' -Value $value ) )
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'proceed' -value ('$null -eq ($addressPool | where startAddress -eq $startAddress)')  ) )
			
			#### Delete Generated range if asked

			$value = 'Get-OVAddressPoolRange| where {($_.name -eq $poolType) -and ($_.rangeCategory -eq "Generated")} '
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'poolToDelete' -Value $value ) )
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '$poolToDelete |  Remove-OVAddressPoolRange -confirm:$False' -isVar $False ) )

		}

		ifBlock			-condition  'if ($proceed)'

		$addressParam 		= ' -Start $startAddress -End $endAddress '
		

		if ($poolType -like "ip*")
		{
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('New-OVAddressPoolRange -IPSubnet $subnet -name "{0}" {1}' -f $name,  $addressParam ) -isVar $False -indentlevel 1))
		}
		else 
		{
			if ($rangeType -eq 'Generated') 
			{
				$addressParam 		= '' 
			}
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('New-OVAddressPoolRange -PoolType $poolType -RangeType $rangeType {0} ' -f  $addressParam ) -isVar $False -indentlevel 1))	
		}
		endIfBlock

		# Skip creating because resource already exists
		elseBlock
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $name + ' already exists.') -isVar $False -indentlevel 2 ))
		endElseBlock


		newLine
	 }
	 
	 if ($List)
	 {
		 [void]$ScriptCode.Add('Disconnect-OVMgmt')
	 }

	# ---------- Generate script to file
	writeToFile -code $ScriptCode -file $ps1Files
 }



 # ---------- Ethernet networks
Function Import-ethernetNetwork([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $scriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	
    foreach ($net in $List)
    {
		$name               = $net.name  
		$type 				= $net.type       
		$vLANType           = $net.ethernetNetworkType
		$vLANID             = $net.vLanId
		$subnetID 			= $net.subnetID
		$ipV6subnetID 		= $net.ipV6subnetID
		$pBandwidth         = (1000 * $net.typicalBandwidth).ToString()
		$mBandwidth         = (1000 * $net.maximumBandwidth).ToString()
		$smartlink          = $net.SmartLink
		$private            = $net.PrivateNetwork
		$purpose            = $net.purpose
		$scopes             = $net.scopes

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating ethernet networks {0} "' -f $name) -isVar $False ))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'net' -Value ("get-OVNetwork | where name -eq '{0}' " -f $name) ))

		ifBlock 		-condition 'if ($Null -eq $net )' 
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('# -------------- Attributes for Ethernet network "{0}"' -f $name) -isVar $False -indentlevel 1))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'name' -Value ('"{0}"' -f $name) -indentlevel 1))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'vLANType' -Value ('"{0}"' -f $vLANType) -indentlevel 1))

		$vLANIDparam = $vLANIDcode = $null

		# --- vLAN
		if ($vLANType -eq 'Tagged')
		{ 

			if (($vLANID) -and ($vLANID -gt 0)) 
			{
				$vLANIDparam = ' -VlanID $VLANID'
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'vLANid' -Value ('{0}' -f $vLANID) -indentlevel 1))

			}

		}                

		# --- Bandwidth
		$pBWparam = $pBWCode = $null
		$mBWparam = $mBWCode = $null

		if ($pBandwidth) 
		{

			$pBWparam = ' -TypicalBandwidth $pBandwidth'
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'pBandwidth' -Value ('{0}' -f $pBandwidth) -indentlevel 1))

		}

		if ($mBandwidth) 
		{

			$mBWparam = ' -MaximumBandwidth $mBandwidth'
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'mBandwidth' -Value ('{0}' -f $mBandwidth) -indentlevel 1))

		}

		# --- subnet
		$subnetCode     = $null
		$subnetIDparam  = ''
		$IPv6subnetCode = $IPv6subnetIDparam = $null
		$subnetArray 	= @()
		
		if ($subnetID)
		{
			
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'ipV4subnet' -Value ("Get-OVAddressPoolSubnet -NetworkID `'{0}`'" -f $subnetID ) -indentlevel 1))
			$subnetArray += '$ipV4subnet'
		}

		if ($ipV6subnetID)
		{
			
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'ipV6subnet' -Value ("Get-OVAddressPoolSubnet -NetworkID `'{0}`'" -f $ipV6subnetID ) -indentlevel 1))
			$subnetArray += '$ipV6subnet'
		}

		if ($subnetArray)
		{
				$value 	= '@({0})' -f ($subnetArray -join $COMMA)
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'subnet' -Value $value -indentlevel 1) )
				$subnetIDparam 	= ' -subnet $subnet'
		}



		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'PLAN' -Value ('${0}' -f $private) -indentlevel 1))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'smartLink' -Value ('${0}' -f $smartLink)-indentlevel 1))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'purpose' -Value ('"{0}"' -f $purpose)-indentlevel 1))

		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix 'New-OVNetwork -Name $name  -PrivateNetwork $PLAN -SmartLink $smartLink -VLANType $VLANType  -purpose $purpose `' -isVar $False -indentLevel 1)) 
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('{0}{1}{2}{3}' -f $vLANIDparam, $pBWparam, $mBWparam, $subnetIDparam) -isVar $False -indentLevel 4))
		newLine # to end the command
		# --- Scopes
		if ($scopes)
		{
			newLine
			[void]$scriptCode.Add( (Generate-CustomVarCode -Prefix 'object' -Value 'get-OVNetwork | where name -eq $name' -indentlevel 1))
			generate-scopeCode -scopes $scopes -indentlevel 1

		}

		endIfBlock 

		# Skip creating because resource already exists
		elseBlock
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $name + ' already exists.') -isVar $False -indentlevel 1 ))
		endElseBlock


		newLine
		

    }

	
	if ($List)
	{
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	}

    # ---------- Generate script to file
    writeToFile -code $ScriptCode -file $ps1Files
    
}


# ---------- networks
Function Import-fcNetwork([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $scriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook 

    foreach ($net in $List)
    {
		$name                   = $net.name
		$type                   = $net.type
		$fabricType             = $net.fabricType
		$managedSan 			= $net.managedSan	
		$vLANID             	= $net.vLanId
		$pBandwidth             = 1000 * $net.typicalBandwidth
		$mBandwidth             = 1000 * $net.maximumBandwidth
		$autoLoginRedistribution = $net.autoLoginRedistribution
		$linkStabilityTime		 = $net.linkStabilityTime
		$scopes 				 = $net.scopes

		if ($type -match 'FC')
		{
			$type 	= 'FibreChannel'
		}

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating FC/FCOE networks {0} "' -f $name) -isVar $False ))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'net' -Value ("get-OVNetwork -type {1} | where name -eq '{0}'  " -f $name, $type) ))

		ifBlock			-condition 'if ($Null -eq $net )' 
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('# -------------- Attributes for FC network "{0}"' -f $name) -isVar $False -indentlevel 1))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'name' -Value ('"{0}"' -f $name) -indentlevel 1))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'type' -Value ('"{0}"' -f $type)-indentlevel 1))

		# --- Bandwidth
		$pBWparam = $pBWCode = $null
		$mBWparam = $mBWCode = $null

		if ($PBandwidth) 
		{

			$pBWparam = ' -typicalBandwidth $pBandwidth'
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'pBandwidth' -Value ('{0}' -f $pBandwidth) -indentlevel 1))

		}

		if ($MBandwidth) 
		{

			$mBWparam = ' -maximumBandwidth $mBandwidth'
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'mBandwidth' -Value ('{0}' -f $mBandwidth) -indentlevel 1))

		}

		# --- ManagedSan
		$SANParam  = $Null
		if ($managedSan)
		{
			$SANparam   = ' -ManagedSAN $managedSAN' 

			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'SANname' -Value ('"{0}"' -f $SANname) -indentlevel 1))
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'managedSAN' -Value ('Get-OVManagedSAN -Name $SANname') -indentlevel 1))		
		}


		# ---- FC or FCOE network
		$FCParam 	 = $linkParam	 =  $autologinParam = $vLanIdParam = $null
		if ($type -eq 'fcoe')
		{
			if (($vLANID) -and ($vLANID -gt 0)) 
			{

				$vLanIdParam  =   ' -vLanID $vLanId'
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'vLanId' -Value ('{0}' -f $vLANID) -indentlevel 1))
			
			}

				
		}
		else # FC network
		{
			$FCparam          = ' -FabricType $fabricType'
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'fabricType' -Value ('"{0}"' -f $fabricType)  -indentlevel 1))
			if ($fabrictype -eq 'FabricAttach')
			{

				if ($autologinredistribution)
				{

					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'autologinredistribution' -Value ('${0}' -f $autologinredistribution)  -indentlevel 1))
					$autologinParam     = ' -AutoLoginRedistribution $autologinredistribution'

				}

				if ($linkStabilityTime) 
				{

					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'LinkStabilityTime' -Value ('{0}' -f $LinkStabilityTime) -indentlevel 1))
					$linkParam  = ' -LinkStabilityTime $LinkStabilityTime'

				}

				$FCparam              += $autologinParam + $linkParam
			}
		}

		# Note : when using backstick (`) make sure that theree is no space after. Otherwise it is considered as escape teh space char and NOT line continuator
		newLine
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('New-OVNetwork -Name $name -Type $Type `') -isVar $False -indentlevel 1) )
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0}{1}{2}{3}{4}' -f $pBWparam, $mBWparam, $FCparam, $vLANIDparam, $SANparam) -isVar $False -indentlevel 4) )
		newLine # to end the command

		# --- Scopes
		if ($scopes)
		{
			newLine
			[void]$scriptCode.Add( (Generate-CustomVarCode -Prefix 'object' -Value 'get-OVNetwork | where name -eq $name' -isVar $False -indentlevel 1))
			generate-scopeCode -scopes $scopes -indentlevel 1

		}

		endIfBlock 

		# Skip creating because resource already exists
		elseBlock
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $name + ' already exists.') -isVar $False -indentlevel 1 ))
		endElseBlock

		newLine


	}

	if ($List)
	{
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	}
    # ---------- Generate script to file
    writeToFile -code $ScriptCode -file $ps1Files

}


# ---------- Network Sets
Function Import-NetworkSet([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
	$scriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  

	

    foreach ($ns in $List)
	{
		$name               = $ns.name
		$PBandwidth         = 1000 * $ns.TypicalBandwidth 
		$Mbandwidth         = 1000 * $ns.MaximumBandwidth 
		$networkSetType 	= $ns.networkSetType
		$networks        	= $ns.networks
		$nativeNetwork 		= $ns.nativeNetwork
		$scopes 			= $ns.scopes

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating networkset {0} "' -f $name) -isVar $False ))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'net' -Value ("get-OVNetworkSet | where name -eq '{0}' " -f $name) ))

		ifBlock			-condition 'if ($Null -eq $net )'  
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('# -------------- Attributes for Network Set "{0}"' -f $name) -isVar $False -indentlevel 1))

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'name' -Value ('"{0}"' -f $name) -indentlevel 1))
			
		$pBWparam = $pbWCode = $null
		$mBWparam = $mBWCode = $null

		if ($PBandwidth) 
		{

			$pBWparam = ' -TypicalBandwidth $pBandwidth'
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'pBandwidth' -Value ('{0}' -f $pBandwidth) -indentlevel 1))

		}
		
		if ($MBandwidth) 
		{

			$mBWparam = ' -MaximumBandwidth $mBandwidth'
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'mBandwidth' -Value ('{0}' -f $mBandwidth) -indentlevel 1))

		}

		# --- networks
		$netParam =  $null
		if ( $networks )
		{
			$netParam     = ' -Networks $networks'

			$arr 		= [system.Collections.ArrayList]::new()
			$arr 		= $networks.split($SepChar) | % { '"{0}"' -f $_ }
			$netList    = '@({0})' -f [string]::Join($Comma, $arr)
			$value 		= ('{0}' -f $netList ) + ' | % { get-OVNetwork -name $_ }'

			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix networks -value $value -indentlevel 1))
		}

		# --- native NEtwork
		$untaggedParam 		= $null

		if ($nativeNetwork)
		{
			$untaggedParam 		= ' -UntaggedNetwork $nativeNetwork'
			$value 				= '"{0}"' -f $nativeNetwork
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix untaggedNetwork -value $value -indentlevel 1))
		}

		newLine
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('New-OVNetworkSet -Name $name{0}{1}{2}{3}' -f $pBWparam, $mBWparam, $netParam, $untaggedParam) -isVar $False -indentlevel 1))

		# --- network Set Type
		#[void]$scriptCode.Add((Generate-CustomVarCode -Prefix nsType  -value ('@{ networkSetType = ' + '"{0}"' -f $networkSetType + "}") -indentlevel 1))
		
		#[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'ns' -value ("Get-OVNetworkSet -name {0}" -f $name) -indentlevel 1))
		##[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'ns.networkSetType' -value '$nsType' -indentlevel 1))
		#[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '$ns.PSobject.Properties.Remove("typicalBandwidth")' -isVar $false -indentlevel 1))
		#[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '$ns.PSobject.Properties.Remove("maximumBandwidth")' -isVar $false -indentlevel 1))
		#[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'Set-OVResource -InputObject $ns | Wait-OVTaskComplete' -isVar $False  -indentlevel 1))
		

		endIfBlock 

        # Skip creating because resource already exists
        elseBlock
        [void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $name + ' already exists.') -isVar $False -indentlevel 1 ))
		endElseBlock
		
		newLine
	}
	
	if ($List)
	{
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	}
    # ---------- Generate script to file
    writeToFile -code $ScriptCode -file $ps1Files

}

# ---------- LIG
Function Import-LogicalInterconnectGroup([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
	$scriptCode             	= [System.Collections.ArrayList]::new()

	$cSheet, $ligSheet, $snmpConfigSheet, $snmpV3UserSheet, $snmpTrapSheet 	= $sheetName.Split($SepChar)       # composer
	$ligList 					= if ($ligSheet)	 		{get-datafromSheet -sheetName $ligSheet -workbook $WorkBook				} else {$null}
	
	#Note: snmp____List will be extracted in teh snmp subsection

	

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	
	$list 						= $ligList 
	foreach ($L in $list)    
	{
		$name  	    				= $L.name
		$FrameCount 		 		= $L.frameCount
		$ICBaySet    				= $L.interConnectBaySet
		$enclosureType				= $L.enclosureType
		$fabricModuleType 			= $L.fabricModuleType
		$bayConfig 					= $L.bayConfig
		$redundancyType  			= $L.redundancyType
		$internalNetworks			= $L.internalNetworks
		$internalNetworkConsistency = if ($L.consistencyCheckingForInternalNetworks) 	{$consistencyCheckingEnum.Item($L.consistencyCheckingForInternalNetworks) } else {'None'}
		$interconnectConsistencyChecking = if ($L.interconnectConsistencyChecking) 		{$consistencyCheckingEnum.Item($L.interconnectConsistencyChecking) } else {'None'}

		$enableIgmpSnooping     	= $L.enableIgmpSnooping
		
		$igmpIdleTimeoutInterval	= $L.igmpIdleTimeoutInterval

		$enableFastMacCacheFailover	= $L.enableFastMacCacheFailover
		$macRefreshInterval			= $L.macRefreshInterval

		$enableNetworkLoopProtection= $L.enableNetworkLoopProtection
		
		$enablePauseFloodProtection	= $L.enablePauseFloodProtection
		$enableRichTLV				= $L.enableRichTLV

		$enableTaggedLldp			= $L.enableTaggedLldp
		$lldpIpAddressMode			= $L.lldpIpAddressMode
		$lldpIpv4Address			= $L.lldpIpv4Address
		$lldpIpv6Address			= $L.lldpIpv6Address

		$enableStormControl			= $L.enableStormControl
		$stormControlPollingInterval = $L.stormControlPollingInterval
		$stormControlThreshold		= $L.stormControlThreshold

		$qosconfigType				= $L.qosconfigType
		$downlinkClassificationType = $L.downlinkClassificationType
		$uplinkClassificationType 	= $L.uplinkClassificationType

		$scopes 					= $L.scopes


		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating logical interconnect group {0} " ' -f $name) -isVar $False ))

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'lig' 						-Value ("get-OVLogicalInterconnectGroup | where name -eq  '{0}' " -f $name) ))

		ifBlock			-condition 'if ($lig -eq $Null)' 
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('# -------------- Attributes for LIG "{0}"' -f $name) -isVar $False -indentlevel 1))

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'name' 						-Value ('"{0}"'	-f $name) -indentlevel 1))
		
		# --- Frame Count - InterconnectBay Set - Fabric Module Type
		$FrameCountParam		= ' -frameCount $frameCount'
		$ICBaySetParam			= ' -interConnectBaySet $interConnectBaySet '
		$fabricModuleParam 		= ' -fabricModuleType $fabricModuleType'
		
		
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'frameCount' 					-Value ('{0}' 	-f $frameCount) -indentlevel 1))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'interconnectBaySet'			-Value ('{0}' 	-f $ICBaySet) -indentlevel 1))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'fabricModuleType' 			-Value ("'{0}'" -f  $fabricModuleType) -indentlevel 1 )) 

		# redundancy Type
		$redundancyParam 	= $null
		if ($redundancyType)
		{
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'redundancyType' 				-Value ('"{0}"'	-f $redundancyType) -indentlevel 1))
			$redundancyParam = ' -FabricRedundancy $redundancyType'
		}
		# Bay Config
			#From the input Frame1={Bay3='SEVC40f8'|Bay6='SEVC40f8'}, we built an array of hash Table
		$baysParam 			= $null
		$isNotSAS 			= $bayConfig -notlike '*SAS*'
		if ($bayConfig)
		{
			$bayConfig 		= [string]($bayConfig.replace($CRLF, ';').replace($CR, ';').replace($sepChar,';') )  # concatenate into string separated with ;
			
			$bayConfig 		= $bayConfig -replace '.$' , '}'   						# Replace last element	

			# Configure IC ConnectType $ICModuleTypes 
			$bayConfig 		= $bayConfig.replace($Space, '')
			foreach ($moduleType in $ICModuleTypes.Keys)
			{
				$bayConfig  = $bayConfig.replace($moduleType, "'{0}'" -f $ICModuleTypes.$moduleType)
			}

			# -- Construct Hash Table
			$bayConfig 		= $bayConfig.replace('={', '=@{')							# for Bay Hash Table
			$value			= '@{' + $bayConfig											# add hash 

			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'bayConfig' -Value  $value -indentlevel 1 ))

			$baysParam 		= ' -Bays $bayConfig'

			

		}

		if ($isNotSAS)
		{
			$igmpParam = $igmpIdleTimeoutParam = $null
			if ($enableIgmpSnooping -eq 'True')
			{
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'enableIgmpSnooping'			-Value ('${0}' 	-f $enableIgmpSnooping) -indentlevel 1))
				if ($igmpIdleTimeoutInterval)
				{
					$igmpIdleTimeoutParam = ' -igmpIdleTimeOutInterval $igmpIdleTimeoutInterval'
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'igmpIdleTimeoutInterval'		-Value ('{0}' 	-f $igmpIdleTimeoutInterval) -indentlevel 1))
				}
				$igmpParam                  = ' -enableIgmpSnooping $enableIgmpSnooping {0}' -f $igmpIdleTimeoutParam
			}


			
			$networkLoopProtectionParam 	= ' -enablenetworkLoopProtection $enableNetworkLoopProtection'
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'enableNetworkLoopProtection'	-Value ('${0}' 	-f $enableNetworkLoopProtection) -indentlevel 1))
			

			$EnhancedLLDPTLVParam       	= ' -enableEnhancedLLDPTLV $enableRichTLV'
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'enableRichTLV'				-Value ('${0}' 	-f $enableRichTLV) -indentlevel 1))

			$LLDPtaggingParam 		      	= ' -enableLLDPTagging $enableTaggedLldp'
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'enableTaggedLldp'			-Value ('${0}' 	-f $enableTaggedLldp) -indentlevel 1))

			$LldpAddressingModeParam		= ' -lldpAddressingMode $lldpIpAddressMode'
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'lldpIpAddressMode'			-Value ("'{0}'" -f $lldpIpAddressMode) -indentlevel 1))			

			#$stormControlParam 				= $null
			#if ($enableStormControl -eq 'True')
			#{
			#	$stormControlParam 				= ' -enableStormControl $enableStormControl '
			#	[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'enableStormControl'			-Value ('${0}' 	-f $enableStormControl) -indentlevel 1))
			#	[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'stormControlPollingInterval'	-Value ('{0}' 	-f $stormControlPollingInterval) -indentlevel 1))
			#	[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'stormControlThreshold'		-Value ('{0}' 	-f $stormControlThreshold) -indentlevel 1))
			#}

			# --- Specific to C7000
			
			if ($enclosureType -ne	$Syn12K)
			{
				$macCacheParam 					= $null
				if ($enableFastMacCacheFailover -eq 'True')
				{
					$macCacheParam 				= ' -enableFastMacCacheFailover $enableFastMacCacheFailover -MacRefreshInterval $macRefreshInterval'
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'enableFastMacCacheFailover'	-Value ('${0}' 	-f $enableFastMacCacheFailover) -indentlevel 1))
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'macRefreshInterval'			-Value ('{0}' 	-f $macRefreshInterval) -indentlevel 1))
				}
				$pauseFloodProtectionParam		= $null
				$pauseFloodProtectionParam 		= ' -enablePauseFloodProtection $enablePauseFloodProtection'
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'enablePauseFloodProtection'	-Value ('${0}' 	-f $enablePauseFloodProtection) -indentlevel 1))

			}

			$InterconnectConsistencyCheckingParam = ' -interconnectConsistencyChecking $interconnectConsistency'
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'interconnectConsistency'		-Value ("'{0}'" 	-f $interconnectConsistencyChecking) -indentlevel 1))

			# ------ Internal Networks
			$intNetParam 	= $null
			if ($internalNetworks)
			{

				$networks 	= $internalNetworks.replace($sepChar, '";"')
				$networks 	= $networks.Insert($networks.length,'")').Insert(0,'@("')
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'internalNetworks'			-Value ('{0}' -f $networks + ' | % {Get-OVNetwork -name $_}' 	) -indentlevel 1))
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'internalNetworkConsistency'	-Value ("'{0}'" -f $internalNetworkConsistency ) -indentlevel 1))				
			
				$intNetParam = ' -InternalNetworks $internalNetworks -internalNetworkConsistencyChecking $internalNetworkConsistency'  
			}

			# ------ snmp
			$snmpParam 		= $null

			$snmpConfigList = if ($snmpConfigSheet)	 		{get-datafromSheet -sheetName $snmpConfigSheet -workbook $WorkBook		} else {$null}
			$snmpConfigList	= $snmpConfigList | where source -eq $name

			if ($snmpConfigList) 		# Have snmpConfiguration for this lig?
			{

				$snmpConfigurationParam =  $communityStringParam  = $contactParam = $snmpV3userParam = $snmpTrapDestinationParam	= $Null
				# snmpV1 information
				$communityString	= $snmpConfigList.communityString
				$contact 			= $snmpConfigList.contact
				$accList 			= $snmpConfigList.accessList

				newLine
				if ($communityString)
				{
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'communityString'	-Value ("'{0}'" -f $communityString ) -indentlevel 1))
					$communityStringParam 	= ' -snmpV1 $True -ReadCommunity $communityString '
				}

				if ($contact) 
				{
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'contact'	-Value ("'{0}'" -f $contact ) -indentlevel 1))
					$contactParam 	= ' -Contact $contact '					
				}
				if ($accList)
				{
					$accList 			= "@(" + ($accList -replace '|', ',') + ")"
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'accessList'	-Value ("{0}" -f $accList ) -indentlevel 1))	
					$accessListParam 	= ' -accessList $accessList ' 
				}

				# ----  snmpV3Users for this list
				$snmpV3UserList 			= if ($snmpV3UserSheet)	 		{get-datafromSheet -sheetName $snmpV3UserSheet -workbook $WorkBook			} else {$null}
				$snmpV3UserList 			= $snmpV3UserList | where source -eq $name
				if ($snmpV3UserList)
				{
					create-snmpV3User -list $snmpV3UserList -isSnmpAppliance $False -indent 1
					$snmpV3userParam 	= ' -snmpV3 $True -snmpV3Users $snmpV3Users '
				}

				# ----- snmpTrap 
				$snmpTrapList 			= if ($snmpTrapSheet)	 		{get-datafromSheet -sheetName $snmpTrapSheet -workbook $WorkBook			} else {$null}
				$snmpTrapList 			= $snmpTrapList | where source -eq $name
				if ($snmpTrapList)
				{
					create-snmpTrap -list $snmpTrapList -isSnmpAppliance $False -indent 1
					$snmpTrapDestinationParam = ' -TrapDestination $snmpTraps  '
				}
				

				# ----- snmpTrapConfiguration 

				$isConfigEmpty = ($Null -eq $communityStringParam) -and ($Null -eq $contactParam) -and ($Null -eq $accessListParam) -and ($Null -eq $snmpV3userParam) -and ($Null -eq $snmpTrapDestinationParam)
				if (-not $isConfigEmpty) 
				{	
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'snmpConfiguration' -value ('New-OVSnmpConfiguration {0}{1}{2}{3}{4} ' -f $communityStringParam, $contactParam, $accessListParam, $snmpV3userParam, $snmpTrapDestinationParam)  -indentlevel 1))
					newLine
					$snmpConfigurationParam 	= ' -snmp $SnmpConfiguration '
				}


			}

			

			$ligVariable    = '$lig'
			newLine
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0} = New-OVLogicalInterconnectGroup -Name $name {1}{2}{3} `' -f $LigVariable, $fabricModuleParam, $FrameCountParam , $ICBaySetParam) -isVar $false -indentlevel 1))
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0}{1} `' 		-f $baysParam, $redundancyParam ) -isVar $false -indentlevel 4))
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0}{1}{2}{3} `' -f $intNetParam,$igmpParam,$pauseFloodProtectionParam, $networkLoopProtectionParam) -isVar $false -indentlevel 4))
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0}{1} `' 		-f $macCacheParam,$EnhancedLLDPTLVParam ) -isVar $false -indentlevel 4))
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0}{1}{2} `' 	-f $LldpAddressingModeParam,$LLDPtaggingParam,$InterconnectConsistencyCheckingParam ) -isVar $false -indentlevel 4))
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0} '			-f $snmpConfigurationParam ) -isVar $false -indentlevel 4))
			
			newLine # to end the command

			#TBD      ,  $snmpParam, $QosParam, $ScopeParam))

		}
		else  # SAS lig
		{
			$ligVariable    = '$lig'
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0} = New-OVLogicalInterconnectGroup -Name $name {1}{2}{3}{4}' -f $LigVariable, $fabricModuleParam, $FrameCountParam , $ICBaySetParam, $baysParam) -isVar $false -indentlevel 1))
				
		}

		endIfBlock -condition 'if ($lig -eq $Null)' 

		# Skip creating because resource already exists
		elseBlock
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $name + ' already exists.') -isVar $False -indentlevel 1 ))
		endElseBlock

		newLine

	}
	
	if ($List)
	{
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	}

	 # ---------- Generate script to file
	 writeToFile -code $ScriptCode -file $ps1Files
}

# ---------- Uplink Set
Function Import-UplinkSet([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $scriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  	
	
	foreach ( $upl in $List)
	{
		$ligName                    = $upl.ligName
        $uplName           			= $upl.name
		$networkType 	            = $upl.networkType
		$networks 			        = $upl.Networks			#[]
		$networkSets				= $upl.NetworkSets		#[]
		$nativeNetwork 		    	= $upl.nativeNetwork
		$enableTrunking 		    = $upl.enableTrunking
		$logicalPortConfigInfos		= $upl.LogicalPortConfigInfos

		$lacpTimer       			= if ($upl.lacpTimer) 			{  $upl.lacpTimer.Trim() } 			else { 'Short' }
		$loadBalancingMode			= $upl.loadBalancingMode
		$primaryPort     			= $upl.PrimaryPort
		$fcSpeed         			= $upl.FCuplinkSpeed #[]
		$consistency 				= if ($upl.consistencyChecking) {$consistencyCheckingEnum.Item($upl.consistencyChecking) } else {'None'}
		
		
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating uplinkset {0} on LIG {1}"' -f $uplName,$ligName) -isVar $False ))
 
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'lig' 						-Value ("get-OVLogicalInterconnectGroup | where name -eq  '{0}' " -f $ligName ) ))		
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'upl' 						-Value ('$lig.uplinksets | where name -eq  "{0}" ' -f $uplName) ))
		newLine

		ifBlock			-condition 'if ( ($lig -ne $Null) -and ($upl -eq $Null) )' 
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('# -------------- Attributes for uplinkset {0} on LIG {1}' -f $uplName,$ligName) -isVar $False -indentlevel 1))

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'name' 						-Value ('"{0}"'	-f $uplName) -indentlevel 1))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'networkType' 				-Value ('"{0}"'	-f $networkType) -indentlevel 1))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'uplConsistency' 				-Value ('"{0}"'	-f $consistency) -indentlevel 1))


		# ---- Networks
		$netParam = $null
		if ($networks)
		{
			$netParam   = ' -Networks $networks'

			$arr 		= [system.Collections.ArrayList]::new()
			$arr 		= $networks.split($SepChar) | % { '"{0}"' -f $_ }
			$netList    = '@({0})' -f [string]::Join($Comma, $arr)
			$value 		= ('{0}' -f $netList ) + ' | % { get-OVNetwork -name $_ }'

			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix networks -value $value -indentlevel 1))
		}


		$netsetParam = $nativeNetParam =  $lacpParam = $trunkingParam = $null
		if ($networkType -ne 'FibreChannel') 
		{
			# ---- Network Sets ( No CopyNetworksFromNetworkSet)
			if ($networkSets)
			{
				$arr 		= [system.Collections.ArrayList]::new()
				$arr 		= $networkSets.split($SepChar) | % { '"{0}"' -f $_ }
				$netList    = '@({0})' -f [string]::Join($Comma, $arr)
				$value 		= ('{0}' -f $netList ) + ' | % { get-OVNetworkSet -name $_ }'

				$netsetParam   = ' -NetworkSets $networkSets'
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix networkSets -value $value -indentlevel 1))
			}

			# ----- Nativenetwork
			if ($nativeNetwork)
			{
				$nativeNetParam = ' -NativeEthNetwork $nativeNetwork'
				$value 			= "get-OVNetwork -name '{0}' " -f $nativeNetwork
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix nativeNetwork -value $value -indentlevel 1))
			}

			# ---- lacpTimer and loadbalancing
			$lacpParam 	= ' -LacpTimer $lacpTimer -LacpLoadbalancingMode $lacpLoadbalancingMode'
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix lacpTimer -value ("'{0}'" -f $lacpTimer) -indentlevel 1))
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix lacpLoadbalancingMode -value ("'{0}'" -f $loadbalancingMode) -indentlevel 1))
		}
		else # Fibre Channel specific
		{
			$trunkingParam 	= ' -enableTrunking $enableTrunking'
			$value 			= '${0}' -f $enableTrunking
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix enableTrunking -value $value -indentlevel 1))

			$trunkingParam 	= ' -fcUplinkSpeed $fcUplinkSpeed'
			$value 			= "'{0}'" -f ($fcSpeed -replace 'Gb','')
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix fcUplinkSpeed  -value $value -indentlevel 1))
		}

		# ---- Logical Ports config -- transform Enclosure1:Bay3:Q1|Enclosure1:Bay3:Q2|Enclosure1:Bay3:Q3 into table
		$uplinkPortParam 	= $null
		if ($logicalPortConfigInfos)
		{
			$uplinkPortParam 	= ' -UplinkPorts $uplinkPorts'
			$value 				= $logicalPortConfigInfos.replace($SepChar, '","') 	# Comma and quote
			$value 				= "@(`"$value`")"
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix uplinkPorts -value $value -indentlevel 1))
		}




		# Make sure there is no space after backtick (`)
		newLine
		[void]$scriptCode.Add((Generate-CustomVarCode -prefix ('New-OVUplinkSet -InputObject $lig -Name $name -Type $networkType `') -isVar $False -indentlevel 1)) 
		[void]$scriptCode.Add((Generate-CustomVarCode -prefix ('{0}{1}{2}{3}{4} `' 	-f $netParam, $netsetParam, $nativeNetParam, $trunkingParam, $lacpParam) -isVar $False -indentlevel 4)) 
		[void]$scriptCode.Add((Generate-CustomVarCode -prefix ('{0} `' 				-f $uplinkPortParam) -isVar $False -indentlevel 4)) 
		[void]$scriptCode.Add((Generate-CustomVarCode -prefix (' -ConsistencyChecking $uplConsistency' ) -isVar $False -indentlevel 4))
		newLine # to end the command
		
		endIfBlock

		# Skip creating because resource already exists
		elseBlock
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW ' + "{0} does not exist or " -f $ligName   + "{0} already exists." -f $uplName ) -isVar $False -indentlevel 1 ))
		endElseBlock

		newLine
        
        
	}
	
	if ($List)
	{
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	}
	
	 # ---------- Generate script to file
	 writeToFile -code $ScriptCode -file $ps1Files



}


Function Import-LogicalSwitchGroup([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $scriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
		
	foreach ( $swg in $List)
	{
	
		$name                   = $swg.name
		$switchType 			= $swg.switchType
		$numberofSwitches 		= $swg.numberofSwitches

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating Logical Switch Group {0} "' -f $name) -isVar $False ))
 
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'swg' 						-Value ("get-OVLogicalSwitchGroup | where name -eq  '{0}' " -f $name ) ))		

		ifBlock			-condition 'if ($swg -eq $Null)' 
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'switchType' 					-Value ('Get-OVSwitchType -name "{0}"'	-f $switchType) -indentlevel 1))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('new-OVLogicalSwitchGroup -name "{0}" -switchType $switchType -NumberOfSwitches {1}' -f $name,$numberofSwitches) -isVar $False -indentlevel 1))
	
		endBlock

		# Skip creating because resource already exists
		elseBlock
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW ' + "{0} does not exist or " -f $name   ) -isVar $False -indentlevel 1 ))
		endElseBlock

		newLine

		# --- Scopes
		if ($scopes)
		{
			[void]$scriptCode.Add( (Generate-CustomVarCode -Prefix 'object' -Value ('Get-OVLogicalSwitchGroup -name -eq "{0}"' -f $name) -indentlevel 1))
			generate-scopeCode -scopes $scopes -indentlevel 1
			newLine

		}
	}
	
	if ($List)
	{
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	}
	
	 # ---------- Generate script to file
	 writeToFile -code $ScriptCode -file $ps1Files


}


Function Import-LogicalSwitch([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $scriptCode             	= [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  

		
	foreach ( $lsw in $List)
	{
	
		$name                   = $lsw.name
		$logicalSwitchGroup		= $lsw.logicalSwitchGroup
		$IsManaged        		= $lsw.managed -eq 'True'
		$switch1Address			= $lsw.switch1Address
		$switch2Address			= $lsw.switch2Address
		$sshUserName 			= $lsw.sshUserName
		$sshPassword 			= $lsw.sshPassword
		$issnmpV3 				= $lsw.snmpType -eq 'snmpV3'
		if ($issnmpV3)
		{
			$snmpV3User 		= $lsw.snmpV3User
			$snmpAuthLevel		= $lsw.snmpAuthLevel	
			$snmpAuthProtocol	= $lsw.snmpAuthProtocol	
			$snmpAuthPassword	= $lsw.snmpAuthPassword	
			$snmpPrivProtocol	= $lsw.snmpPrivProtocol	
			$snmpPrivPassword	= $lsw.snmpPrivPassword
		}
		else 
		{
			$snmpCommunity 		= $lsw.snmpCommunity
			$snmpPort 			= $lsw.snmpPort	
		}

			

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating Logical Switch {0} "' -f $name) -isVar $False ))
 
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'lsw' 						-Value ("get-OVLogicalSwitch | where name -eq  '{0}' " -f $name ) ))		

		ifBlock			-condition 'if ($lsw -eq $Null)' 
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'logicalSwitchGroup' 			-Value ('Get-OVLogicalSwitchGroup -name "{0}"'	-f $logicalSwitchGroup) -indentlevel 1))
		
		$namegroupParam = ' -name "{0}" -logicalSwitchGroup $logicalSwitchGroup ' -f $name
		$managedParam 	= if ($isManaged) { ' -Managed'} else {' -Monitored'}

		$s1 			= if ($switch1Address) { ' -Switch1Address {0} ' -f $switch1Address} else {''}
		$s2 			= if ($switch2Address) { ' -Switch2Address {0} ' -f $switch2Address} else {''}
		$addressParam 	= $s1 + $s2

		generate-credentialCode -password $sshPassword -username $sshUserName -component 'LOGICAL SWITCH'-indentLevel 1 -scriptCode $scriptCode
		$credParam 		= ' -sshUserName "{0}" -sshPassword $securePassword' -f $sshUserName # $securePassword is defined in  generate-credentialCode

		 
		if ($issnmpV3)
		{
			$snmpParam 	= ' -snmpV3 $True -SnmpUserName "{0}" -SnmpAuthLevel "{1}" ' -f $snmpV3User, $snmpAuthLevel
			if ($snmpAuthLevel -eq 'Auth')
			{	
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'authPassword' -value ('"{0}" | ConvertTo-SecureString -AsPlainText -Force' -f $snmpAuthPassword) -indentLevel 1 ))
				$snmpParam	+= ' -snmpAuthProtocol "{0}" -snmpAuthPassword $authPassword ' -f  $snmpAuthProtocol
			}
			else
			{
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'authPassword' -value ('"{0}" | ConvertTo-SecureString -AsPlainText -Force' -f $snmpAuthPassword) -indentLevel 1 ))
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'privPassword' -value ('"{0}" | ConvertTo-SecureString -AsPlainText -Force' -f $snmpPrivPassword)  -indentLevel 1 ))
				$snmpParam	+=  ' -snmpAuthProtocol "{1}" -snmpAuthPassword $authPassword -snmpPrivProtocol "{2}" -snmpPrivPassword $privPassword' -f $snmpAuthProtocol, $snmpPrivProtocol			
			}
		}
		else # snmpV1
		{
			$snmpParam 	= ' -snmpV1 $True -snmpPort {0} -snmpCommunity "{1}" ' -f $snmpPort, $snmpCommunity 
		}
		
		
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('new-OVLogicalSwitch {0}{1} `' -f $namegroupParam, $managedParam) -isVar $False  -indentlevel 1))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0}{1} `' 						-f $addressParam, $credParam) 	   -isVar $False  -indentlevel 4))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0}' 							-f $snmpParam) 					   -isVar $False  -indentlevel 4))
		newLine
	
		endBlock

		# Skip creating because resource already exists
		elseBlock
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW ' + "{0} does not exist or " -f $name   ) -isVar $False -indentlevel 1 ))
		endElseBlock

		newLine

	}
	
	if ($List)
	{
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	}
	
	 # ---------- Generate script to file
	 writeToFile -code $ScriptCode -file $ps1Files


}


Function Import-SANmanager([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $scriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  

		
	foreach ( $san in $List)
	{
	
		$name                   = $san.name
		$type 					= $san.type
		$userName 				= $san.userName
		$password 				= $san.password
		$useSSL         		= $san.useSSL -eq 'True'
		$port 					= $san.port	

		$snmpUserName 			= $san.snmpUserName
		$snmpAuthLevel			= $san.snmpAuthLevel	
		$snmpAuthProtocol		= $san.snmpAuthProtocol	
		$snmpAuthPassword		= $san.snmpAuthPassword	
		$snmpPrivProtocol		= $san.snmpPrivProtocol	
		$snmpPrivPassword		= $san.snmpPrivPassword


			

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating SAN Manager {0} "' -f $name) -isVar $False ))
 
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'san' 						-Value ("get-OVSANManager | where name -eq  '{0}' " -f $name ) ))		
		
		ifBlock			-condition 'if ($san -eq $Null)' 		
		$nameParam 		= ' -HostName "{0}" -type "{1}" ' -f $name, $type
		if ( ($userName) -and ($password) )
		{
			generate-credentialCode -password $password -username $userName -component 'SAN MANAGER'-indentLevel 1 -scriptCode $scriptCode
			$credParam 		= ' -Credential $cred'  		# $cred is defined in  generate-credentialCode
			$useSSLParam 	= if ($useSSL) 	{ ' -useSSL'            } 	else {''}
			$portParam 		= if ($port)	{' -Port {0} ' -f $port } 	else {''}
			$authParam 		= $credParam + $portParam + $useSSLParam
		}
		else 
		{
			$authParam 	= '  -SnmpUserName "{0}" -SnmpAuthLevel "{1}" ' -f $snmpUserName, $snmpAuthLevel
			switch ($snmpAuthLevel)
			{
				'AuthOnly'
					{	
						[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'authPassword' -value ('"{0}" | ConvertTo-SecureString -AsPlainText -Force' -f $snmpAuthPassword) -indentLevel 1 ))
						$authParam	+= ' -snmpAuthProtocol "{0}" -snmpAuthPassword $authPassword ' -f  $snmpAuthProtocol
					}
				'AuthAndPriv'
					{
						[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'authPassword' -value ('"{0}" | ConvertTo-SecureString -AsPlainText -Force' -f $snmpAuthPassword) -indentLevel 1 ))
						[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'privPassword' -value ('"{0}" | ConvertTo-SecureString -AsPlainText -Force' -f $snmpPrivPassword)  -indentLevel 1 ))
						$authParam	+=  ' -snmpAuthProtocol "{1}" -snmpAuthPassword $authPassword -snmpPrivProtocol "{2}" -snmpPrivPassword $privPassword' -f $snmpAuthProtocol, $snmpPrivProtocol			
					}
			}	
		}

		
		
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('new-OVSANManager {0} `' 	-f $nameParam) -isVar $False  -indentlevel 1))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0}' 						-f $authParam) -isVar $False  -indentlevel 3))
		newLine
	
		endBlock

		# Skip creating because resource already exists
		elseBlock
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW ' + "{0} does not exist." -f $name   ) -isVar $False -indentlevel 1 ))
		endElseBlock

		newLine

	}
	
	if ($List)
	{
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	}
	
	 # ---------- Generate script to file
	 writeToFile -code $ScriptCode -file $ps1Files


}



Function Import-StorageSystem([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $scriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
		
	foreach ( $sts in $List)
	{
	
		$name                   = $sts.name
		$family 				= $sts.family
		$domain 				= $sts.domain
		$userName 				= $sts.userName
		$password 				= $sts.password
		$showSystemDetails      = $sts.showSystemDetails -eq 'True'
		$vips 					= $sts.vips
		$storagePool 			= $sts.StoragePool



			

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating Storage System {0} "' -f $name) -isVar $False ))
 
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'sts' 						-Value ("get-OVStorageSystem | where name -eq  '{0}' " -f $name ) ))		
		ifBlock 'if ($sts -eq $Null)' 	
		
		$s 					= if ($showSystemDetails) { ' -ShowSystemDetails'} else {''}
		$nameParam 			= (' -HostName "{0}" -Family "{1}" ' -f $name, $family) + $s


		$authParam 			= $null
		if ( ($userName) -and ($password) )
		{
			generate-credentialCode -password $password -username $userName -component 'STORAGE SYSTEM'-indentLevel 1 -scriptCode $scriptCode
			$authParam 		= ' -Credential $cred'  		# $cred is defined in  generate-credentialCode
		}

		if ($vips)
		{
			$ip , $netName 	= $vips.Split($Equal)
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'net' 	-Value ("get-OVNetwork | where name -eq  '{0}' " -f $netName.trim() ) -indentlevel 1 ))	
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'vips' 	-Value ('@{' + ('"{0}" = $net' -f $ip.trim())  + '}') 							-indentlevel 1 )) 
			$vipsParam 			 = ' -VIPS $vips'
		}
		
		
		$domainParam 		= if ($domain) { ' -Domain "{0}" ' -f $domain } else {''}
		
		
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('new-OVStorageSystem {0}{1} `' 	-f $nameParam, $authParam) 	-isVar $False  -indentlevel 1))

		if ($family -eq 'StoreServ')
		{
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0} | wait-OVTaskComplete' 	-f $domainParam) 			-isVar $False  -indentlevel 3))

		}
		else
		{
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0} | wait-OVTaskComplete'	-f $vipsParam) 				-isVar $False  -indentlevel 3))
		}
		newLine

		if ($storagePool)
		{
			$pool 		= "@('" + $storagePool.replace($sepChar, "'$comma'") + "')"
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'pool' 						-Value ("{0}" -f $pool ) -indentlevel 1 ))	
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'sts' 						-Value ("get-OVStorageSystem | where name -eq  '{0}' " -f $name ) -indentlevel 1 ))	
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'new-OVStoragePool -pool $pool -StorageSystem $sts | wait-OVTaskComplete'  -isVar $False  -indentlevel 1))
			# Set pool to managed state
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '$pool | % {get-OVStoragePool -name $_ | set-OVStoragePool -Managed $true }'  -isVar $False  -indentlevel 1))
			
			newLine
		}

		endIfBlock
		# Skip creating because resource already exists
		elseBlock
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW ' + "{0} does not exist." -f $name   ) -isVar $False -indentlevel 1 ))
		endElseBlock

		newLine

	}
	
	if ($List)
	{
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	}
	
	 # ---------- Generate script to file
	 writeToFile -code $ScriptCode -file $ps1Files


}


Function set-Param([string]$var, [string]$value)
{
	$param 				= if ($value -eq 'True') { " -$var "} else {""}
	return $param

}
Function Import-StorageVolumeTemplate([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $scriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
		
	foreach ( $svt in $List)
	{
		$name                   			= $svt.name
		$description 						= $svt.description 

		$storagePool 						= $svt.StoragePool
		$lockStoragePool					= set-Param -var 'LockStoragePool' 				-value $svt.lockStoragePool 
		$snapshotStoragePool 				= $svt.snapshotStoragePool
		$lockSnapshotStoragePool			= set-Param -var 'LockSnapshotStoragePool' 		-value $svt.lockSnapshotStoragePool 		
		$storageSystem 						= $svt.StorageSystem
		
		$provisioningType					= $svt.provisioningType
		$lockProvisionType					= set-Param -var 'lockProvisionType'			-value $svt.lockProvisionType
		$capacity 							= if ($svt.capacity) {$svt.capacity} else {1}     # default value is 1 GiB 
		$lockCapacity						= set-Param -var 'LockCapacity' 				-value $svt.lockCapacity
		$shared 							= set-Param -var 'Shared'						-value $svt.shared
		$lockProvisionMode 					= set-Param -var 'LockProvisionMode'			-value ($svt.lockProvisionMode) 

		$enableAdaptiveOptimization 		= set-Param -var 'EnableAdaptiveOptimization' 	-value $svt.enableAdaptiveOptimization
		$lockAdaptiveOptimization			= set-Param -var 'LockAdaptiveOptimization' 	-value $svt.lockAdaptiveOptimization		
		$cachePinning						= set-Param -var 'CachePinning'					-value $svt.cachePinning  
		$lockCachePinning					= set-Param -var 'LockCachePinning'				-value $svt.lockCachePinning
		$dataTransferLimit					= $svt.dataTransferLimit
		$lockDataTransferLimit				= set-Param -var 'LockDataTransferLimit'		-value $svt.lockDataTransferLimit
		$enableDeduplication				= set-Param -var 'EnableDeduplication'			-value $svt.enableDeduplication
		$lockEnableDeduplication			= set-Param -var 'LockEnableDeduplication' 		-value $svt.lockEnableDeduplication
		$enableEncryption					= set-Param -var 'EnableEncryption' 			-value $svt.enableEncryption
		$lockEnableEncryption				= set-Param -var 'LockEnableEncryption' 		-value $svt.lockEnableEncryption
		$folder 							= $svt.folder
		$lockFolder							= set-Param -var 'LockFolder' 					-value $svt.lockFolder
		$IOPSLimit 							= $svt.IOPSLimit
		$lockIOPSLimit 						= set-Param -var 'LockIOPSLimit'				-value $svt.lockIOPSLimit
		$performancePolicy					= $svt.performancePolicy
		$lockPerformancePolicy				= set-Param -var 'LockPerformancePolicy' 		-value $svt.lockPerformancePolicy
		$dataProtectionLevel				= $svt.dataProtectionLevel
		$lockProtectionLevel				= set-Param -var 'LockProtectionLevel'			-value $svt.lockProtectionLevel
		$volumeSet 							= $svt.volumeSet
		$lockVolumeSet 						= set-Param -var 'LockVolumeSet'				-value $svt.lockVolumeSet

		$scopes 							= $svt.scopes
			

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating Storage Volume Template {0} "' -f $name) -isVar $False ))
 
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'svt' 						-Value ("get-OVStorageVolumeTemplate | where name -eq  '{0}' " -f $name ) ))		
		
		ifBlock			-condition 'if ($svt -eq $Null)' 		
		$descParam 			= if ($description) { ' -Description "{0}" ' -f $description} else {''}
		$nameParam 			= (' -Name "{0}" {1} ' -f $name, $descParam) 

		if ($storagePool)
		{	
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'stp' 						-Value ("get-OVStoragePool | where name -eq  '{0}' " -f $storagePool ) -indentlevel 1 ))	
			
			ifBlock		-condition 'if ($stp -ne $Null)' 	-indentlevel 1
			# ---- Storage Pool and SnapshotStoragePool
			$storagePoolParam 			= ' -StoragePool $stp ' + $lockStoragePool
			if ($snapshotStoragePool)
			{
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'sstp' 					-Value ("get-OVStoragePool | where name -eq  '{0}' " -f $snapshotStoragePool ) -indentlevel 2 ))	
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'sstp' -value 'if ($sstp -ne $Null) { $sstp } else {$stp}'  -indentlevel 2) )
				$storagePoolParam 		+= ' -SnapshotStoragePool $sstp ' + $lockSnapshotStoragePool
			}

			# ---- Storage System
			if ($storageSystem)		
			{
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ' # ----- Get Storage System' -isVar $False -indentlevel 2 ))
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'sts' 						-Value ("get-OVStorageSystem | where name -eq  '{0}' " -f $storageSystem ) -indentlevel 2 ))	
				$storagePoolParam 		+= ' -storageSystem $sts ' 
			}


			# ---- Provisioning and Capacity
			$capacityParam 				= (' -Capacity {0} ' -f $capacity) + $lockCapacity

			# Attributes
			$a1							= if ($provisioningType) 			{ (' -ProvisioningType "{0}" ' -f $provisioningType) + $lockProvisionType 	}		else {''}
			$a2 						= if ($enableAdaptiveOptimization)	{ $enableAdaptiveOptimization  + $lockAdaptiveOptimization					}		else {''}
			$a3 						= if ($shared)						{ $shared + $lockProvisionMode 												}		else {''}
			$a4 						= if ($cachePinning)				{ $cachePinning  + $lockCachePinning										}		else {''}			
			$attributes1Param			= $a1 + $a2 + $a3 + $a4

			$a5 						= if ($dataTransferLimit)			{ (' -DataTransferLimit {0}' -f $dataTransferLimit)  + $lockDataTransferLimit}		else {''}
			$a6 						= if ($enableCompression)			{ $enableCompression + $lockEnableEncryption								}   	else {''}
			$a7							= if ($enableDeduplication)			{ $enableDeduplication + $lockEnableDeduplication							}		else {''}
			$a8 						= if ($enableEncryption)			{ $enableEncryption + $lockEnableEncryption									}		else {''}
			$attributes2Param			= $a5 + $a6 + $a7 + $a8

			$a9 						= if ($IOPSLimit)					{ (' -IOPSLimit {0}' -f $IOPSLimit) + $lockIOPSLimit						}		else {''}
			$a10 						= if ($dataProtectionLevel)			{ (' -DataProtectionLevel "{0}" ' -f $dataProtectionLevel) + $lockProtectionLevel}	else {''}

			$attributes3Param			= $a9 + $a10 

			$a12 						= $null
			if ($folder)
			{
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'folder' 	-Value ("(get-OVStoragePool -name default).deviceSpecificAttributes.Folders | where name -eq  '{0}' " -f $folder ) -indentlevel 2 ))	
				$a12 					= ' -Folder $folder' + $lockFolder													
			}

			$a13						= $null 
			if ($performancePolicy)
			{
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'performancePolicy' 	-Value ('Show-OVStorageSystemPerformancePolicy -InputObject $sts -name "{0}" ' -f $performancePolicy ) -indentlevel 2 ))	
				$a13 					= ' -PerformancePolicy $performancePolicy' + $lockPerformancePolicy	
			}
			
			$a14						= $null 
			if ($volumeSet)
			{
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'volumeSet' 	-Value ("get-OVStorageVolumeSet | where name -eq  '{0}' " -f $volumeSet ) -indentlevel 2 ))	
				$a14 					= ' -VolumeSet $volumeSet' + $lockVolumeSet
			}
			$attributes4Param			= $a12 + $a13 + $a14

			#---- code here
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('new-OVStorageVolumeTemplate {0}{1}{2} `'	-f $nameParam, $capacityParam, $storagePoolParam )	-isVar $False  -indentlevel 2))
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0}{1} `'									-f $attributes1Param, $attributes2Param) 			-isVar $False  -indentlevel 3))
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0} `'										-f $attributes3Param) 								-isVar $False  -indentlevel 3))
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0}'										-f $attributes4Param) 								-isVar $False  -indentlevel 3))
			
			newLine

			# --- Scopes
			if ($scopes)
			{
				newLine
				[void]$scriptCode.Add( (Generate-CustomVarCode -Prefix 'object' -Value ('getHPOVStorageVolumeTemplate | where name -eq "{0}"' -f $name) -indentlevel 1))
				generate-scopeCode -scopes $scopes -indentlevel 1

			}

			endifBlock 		-condition 'if ($stp -ne $Null)'  -indentlevel 1

			elseBlock 	-indentlevel 1
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW "storage pool {0} does not exist. Cannot create volume template"' -f $storagepool) 		-isVar $False -indentlevel 2) )
			newLine
			endElseBlock -indentlevel 1


			endIfBlock  -condition 'if $svt -eq $Null'
			elseBlock
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW "volume template {0} already exists. skip creating volume template" ' -f $name) 	-isVar $False -indentlevel 1) )	
			endElseBlock



		}
		else {}#TBD 



	}
	
	if ($List)
	{
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	}
	
	 # ---------- Generate script to file
	 writeToFile -code $ScriptCode -file $ps1Files


}


Function Import-StorageVolume([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $scriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
		
	foreach ( $vol in $List)
	{
		$name                   			= $vol.name
		$description 						= $vol.description 
		$volumeTemplate 					= $vol.volumeTemplate
		
		$storagePool 						= $vol.StoragePool
		$snapshotStoragePool 				= $vol.snapshotStoragePool		
		$storageSystem 						= $vol.StorageSystem

		$provisioningType					= $vol.provisioningType
		$capacity 							= if ($vol.capacity) {$vol.capacity} else {1}     # default value is 1 GiB 

		$enableAdaptiveOptimization 		= set-Param -var 'EnableAdaptiveOptimization' 	-value $vol.enableAdaptiveOptimization		
		$cachePinning						= set-Param -var 'CachePinning'					-value $vol.cachePinning  
		$dataTransferLimit					= $vol.dataTransferLimit
		$enableDeduplication				= set-Param -var 'EnableDeduplication'			-value $vol.enableDeduplication
		$enableEncryption					= set-Param -var 'EnableEncryption' 			-value $vol.enableEncryption
		$folder 							= $vol.folder
		$IOPSLimit 							= $vol.IOPSLimit
		$performancePolicy					= $vol.performancePolicy
		$dataProtectionLevel				= $vol.dataProtectionLevel
		$shared 							= set-Param -var 'Shared'						-value $vol.shared
		$volumeSet 							= $vol.volumeSet

		$scopes 							= $vol.scopes
			

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating Storage Volume {0} "' -f $name) -isVar $False ))
 
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'vol' 						-Value ("get-OVStorageVolume | where name -eq  '{0}' " -f $name ) ))		
		
		ifBlock 		-condition  'if ($Null -eq $vol)' 		
		$descParam 				= if ($description) { ' -Description "{0}" ' -f $description} else {''}
		$nameParam 				= (' -Name "{0}" {1} ' -f $name, $descParam) 

		if ($volumeTemplate)
		{
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'svt' 					-Value ("get-OVStorageVolumeTemplate | where name -eq  '{0}' " -f $volumeTemplate ) -indentlevel 1 ))
			ifBlock -condition  'if ($Null -ne $svt)' -indentlevel 1
			$volumeTemplateParam 	= ' -VolumeTemplate $svt'

			#---- code here
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('new-OVStorageVolume {0}{1}'	-f $nameParam, $volumeTemplateParam )	-isVar $False  -indentlevel 2))
			
			newLine

			# --- Scopes
			if ($scopes)
			{
				newLine
				[void]$scriptCode.Add( (Generate-CustomVarCode -Prefix 'object' -Value ('get-OVStorageVolume | where name -eq "{0}"' -f $name) -indentlevel 1))
				generate-scopeCode -scopes $scopes -indentlevel 1

			}
			endIfBlock  -indentlevel 1
			elseBlock   -indentlevel 1
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW "volume template {0} not found. skip creating volume" ' -f $volumeTemplate) 	-isVar $False -indentlevel 2) )
			endElseBlock -indentlevel 1
			
		}
		else # standalone volume no template
		{
			if ($storagePool)
			{	
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'stp' 						-Value ("get-OVStoragePool | where name -eq  '{0}' " -f $storagePool ) -indentlevel 1 ))	
				ifBlock -condition 'if ( $Null -ne $stp)' 		-isVar $False -indentlevel 1

				# ---- Storage Pool and SnapshotStoragePool
				$storagePoolParam 			= ' -StoragePool $stp ' + $lockStoragePool
				if ($snapshotStoragePool)
				{
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'sstp' 					-Value ("get-OVStoragePool | where name -eq  '{0}' " -f $snapshotStoragePool ) -indentlevel 2 ))	
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'sstp' -value 'if ($sstp -ne $Null) { $sstp } else {$stp}'  -indentlevel 2) )
					$storagePoolParam 		+= ' -SnapshotStoragePool $sstp ' + $lockSnapshotStoragePool
				}

				# ---- Storage System
				if ($storageSystem)		
				{
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ' # ----- Get Storage System' -isVar $False -indentlevel 2 ))
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'sts' 						-Value ("get-OVStorageSystem | where name -eq  '{0}' " -f $storageSystem ) -indentlevel 2 ))	
					$storagePoolParam 		+= ' -storageSystem $sts ' 
				}


				# ---- Provisioning and Capacity
				$capacityParam 				= (' -Capacity {0} ' -f $capacity) + $lockCapacity

				# Attributes
				$a1							= if ($provisioningType) 			{ (' -ProvisioningType "{0}" ' -f $provisioningType) + $lockProvisionType 	}		else {''}
				$a2 						= if ($enableAdaptiveOptimization)	{ $enableAdaptiveOptimization  + $lockAdaptiveOptimization					}		else {''}
				$a3 						= if ($shared)						{ $shared + $lockProvisionMode 												}		else {''}
				$a4 						= if ($cachePinning)				{ $cachePinning  + $lockCachePinning										}		else {''}			
				$attributes1Param			= $a1 + $a2 + $a3 + $a4

				$a5 						= if ($dataTransferLimit)			{ (' -DataTransferLimit {0}' -f $dataTransferLimit)  + $lockDataTransferLimit}		else {''}
				$a6 						= if ($enableCompression)			{ $enableCompression + $lockEnableEncryption								}   	else {''}
				$a7							= if ($enableDeduplication)			{ $enableDeduplication + $lockEnableDeduplication							}		else {''}
				$a8 						= if ($enableEncryption)			{ $enableEncryption + $lockEnableEncryption									}		else {''}
				$attributes2Param			= $a5 + $a6 + $a7 + $a8

				$a9 						= if ($IOPSLimit)					{ (' -IOPSLimit {0}' -f $IOPSLimit) + $lockIOPSLimit						}		else {''}
				$a10 						= if ($dataProtectionLevel)			{ (' -DataProtectionLevel "{0}" ' -f $dataProtectionLevel) + $lockProtectionLevel}	else {''}

				$attributes3Param			= $a9 + $a10 

				$a12 						= $null
				if ($folder)
				{
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'folder' 	-Value ("(get-OVStoragePool -name default).deviceSpecificAttributes.Folders | where name -eq  '{0}' " -f $folder ) -indentlevel 2 ))	
					$a12 					= ' -Folder $folder' + $lockFolder													
				}

				$a13						= $null 
				if ($performancePolicy)
				{
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'performancePolicy' 	-Value ('Show-OVStorageSystemPerformancePolicy -InputObject $sts -name "{0}" ' -f $performancePolicy ) -indentlevel 2 ))	
					$a13 					= ' -PerformancePolicy $performancePolicy' + $lockPerformancePolicy	
				}
				
				$a14						= $null 
				if ($volumeSet)
				{
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'volumeSet' 	-Value ("get-OVStorageVolumeSet | where name -eq  '{0}' " -f $volumeSet ) -indentlevel 2 ))	
					$a14 					= ' -VolumeSet $volumeSet' + $lockVolumeSet
				}
				$attributes4Param			= $a12 + $a13 + $a14

				#---- code here
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('new-OVStorageVolume {0}{1}{2} `'	-f $nameParam, $capacityParam, $storagePoolParam )	-isVar $False  -indentlevel 2))
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0}{1} `'									-f $attributes1Param, $attributes2Param) 			-isVar $False  -indentlevel 3))
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0} `'										-f $attributes3Param) 								-isVar $False  -indentlevel 3))
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0}'										-f $attributes4Param) 								-isVar $False  -indentlevel 3))
				
				newLine

				# --- Scopes
				if ($scopes)
				{
					newLine
					[void]$scriptCode.Add( (Generate-CustomVarCode -Prefix 'object' -Value ('get-OVStorageVolume | where name -eq "{0}"' -f $name) -indentlevel 1))
					generate-scopeCode -scopes $scopes -indentlevel 1

				}

				endBlock 	-indentlevel 1

				elseBlock 	-indentlevel 1
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW "storage pool {0} does not exist. Cannot create volume template"' -f $storagepool) 		-isVar $False -indentlevel 2) )
				newLine
				endElseBlock -indentlevel 1

			}
			else {}#TBD 
		}

		endIfBlock # end check on if $vol -eq $Null

		elseBlock
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW "volume {0} already exists. skip creating volume" ' -f $name) 	-isVar $False -indentlevel 1) )
		endElseBlock




	}
	
	if ($List)
	{
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	}
	
	 # ---------- Generate script to file
	 writeToFile -code $ScriptCode -file $ps1Files


}


Function Import-LogicalJBOD([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $scriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
		
	foreach ( $jbod in $List)
	{
		$name                   			= $jbod.name
		$description 						= $jbod.description 

		$numberofDrives						= $jbod.numberofDrives	
		$driveType							= $jbod.driveType
		$maxDriveSize						= $jbod.maxDriveSize
		$minDriveSize						= if ([string]::IsNullOrEmpty($jbod.minDriveSize)) {$jbod.minDriveSize} else {1}
		$driveEnclosure						= $jbod.driveEnclosure	
		$eraseDataOnDelete					= $jbod.eraseDataOnDelete					

		$scopes 							= $jbod.scopes



		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '# ----------------------------------------------------------------'  -isVar $False ))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating Logical JBOD {0} "' -f $name) -isVar $False ))
 
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'jbod' 						-Value ("get-OVLogicalJBOD| where name -eq  '{0}' " -f $name ) ))		
		
		ifBlock 		-condition	'if ($Null -eq $jbod )' 		
		$descParam 				= if ($description) { ' -Description "{0}" ' -f $description} else {''}
		$nameParam 				= (' -Name "{0}" {1} ' -f $name, $descParam) 
		$eraseParam 			= if ($eraseDataOnDelete -eq 'True') { ' -eraseDataOnDelete $True'} else {''}

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'driveEnclosure' 				-Value ("Get-OVDriveEnclosure | where name -eq  '{0}' " -f $driveEnclosure) -indentlevel 1 ))
		
		ifBlock 		-condition 'if ( $Null -ne $driveEnclosure )'  -indentlevel 1
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('New-OVLogicalJbod	{0} `' 									-f $nameParam) 										-isVar $False -indentlevel 2)) 
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix  ' -InputObject $driveEnclosure `'  																				-isVar $False -indentlevel 7)) 
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix (' -NumberofDrives {0} -MinDriveSize {1} -MaxDriveSize {2} `'	-f $numberofDrives, $minDriveSize , $MaxDriveSize)	-isVar $False -indentlevel 7)) 
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix (' -DriveType {0} {1} '										-f $driveType, $eraseParam ) 						-isVar $False -indentlevel 7)) 
			
		endIfBlock   	-indentlevel 1 # end of check drive enclosure null

		elseBlock 		-indentlevel 1
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW "No such drive enclosure {0} to create JBOD."' -f $driveEnclosure)	-isVar $False -indentlevel 2) )
		endElseBlock 	-indentlevel 1

		endIfBlock 		-condition 'if $Null -eq $jbod '
		elseBlock
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW "logical JBOD {0} already exists. skip creating JBOD" ' -f $name) 	-isVar $False -indentlevel 1) )
		endElseBlock



	}



	
	if ($List)
	{
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	}

	# ---------- Generate script to file
	writeToFile -code $ScriptCode -file $ps1Files


}


# ---------- Enclosure Group
Function Import-EnclosureGroup([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $scriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
		
	foreach ( $eg in $List)
	{
	
		$name                   = $eg.name
        $ligMapping             = $eg.logicalInterConnectGroupMapping
		$enclosureCount         = $eg.enclosureCount
		$ipV4AddressingMode		= $eg.ipV4AddressingMode
		$ipV4Range 				= $eg.ipV4Range #[]
		$ipV6AddressingMode		= $eg.ipV6AddressingMode
		$ipV6Range 				= $eg.ipV6Range #[]
		$powerMode              = $eg.powerMode
		$deploymentMode			= $eg.deploymentMode
		$deploymentNetwork		= $eg.deploymentNetwork
        $scopes              	= $eg.scopes

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating enclosure Group {0} "' -f $name) -isVar $False ))
 
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'eg' 						-Value ("get-OVEnclosureGroup | where name -eq  '{0}' " -f $name ) ))		

		ifBlock			-condition 'if ($Null -eq $eg)' 
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('# -------------- Attributes for enclosure group {0} ' -f $name) -isVar $False -indentlevel 1))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'name' 						-Value ('"{0}"'	-f $name) -indentlevel 1))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'enclosureCount' 				-Value ('{0}'	-f $enclosureCount) -indentlevel 1))

		# --- IP v4 addressing Mode
		$v4AddressPoolParam = $null
		if ([string]::IsNullOrWhiteSpace($IPv4AddressingMode) )
		{
			$ipV4AddressingMode = 'DHCP'
		}
		$ipV4AddressingMode			= $ipV4AddressingMode.Trim()
		$v4AddressPoolParam 		=   ' -IPv4AddressType $ipV4AddressType'
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'ipV4AddressType' 			-Value ('"{0}"'	-f $ipV4AddressingMode) -indentlevel 1))

		
		if ($ipV4AddressingMode -eq 'AddressPool')
		{
			$v4Range 				= 	"@('" + $ipV4Range.replace($SepChar, "','") + "')"
			$value 					= $v4Range + ' | % {Get-OVAddressPoolRange | where name -eq $_ } '
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'IPv4AddressRange' 		-Value $value -indentlevel 1))

			$v4AddressPoolParam 	+=   ' -IPv4AddressRange $IPv4AddressRange'
		}

		# --- IP v6 addressing Mode
		$v6AddressPoolParam = $null
		if ($ipV6AddressingMode)
		{
			$ipV6AddressingMode			= $ipV6AddressingMode.Trim()
			$v6AddressPoolParam 		=   ' -IPv6AddressType $ipV6AddressType'
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'ipV6AddressType' 			-Value ('"{0}"'	-f $ipV6AddressingMode) -indentlevel 1))

			
			if ($ipV6AddressingMode -eq 'AddressPool')
			{
				$v6Range 				= 	"@('" + $ipV6Range.replace($SepChar, "','") + "')"
				$value 					= $v6Range + ' | % {Get-OVAddressPoolRange | where name -eq $_  } '
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'IPv6AddressRange' 		-Value $value -indentlevel 1))
	
				$v6AddressPoolParam 	+=   ' -IPv6AddressRange $IPv6AddressRange'
			}
		}

		# Power Mode
		if ($powerMode)
		{
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'powerMode' -Value ('"{0}"' -f $powerMode) -indentlevel 1))
			$powerModeParam         = ' -PowerRedundantMode $powerMode'
		}

		# LIGMapping
		if ($ligMapping)
		{
			$ligGroupMapping 	= $ligMapping -replace '\s+=\s+','=$' -replace ',', ',$'     # Add $ in front of LigName

			# 1 - Build variables
			$vars 				= $ligMapping  -replace 'Frame\d+\s+=\s+' , ''			# Remove Framex=
			$vars 				= $vars.replace($SepChar, $Comma)
			$varArray			= $vars.Split($Comma)								# Build array of variable names
			$varArray 			= $varArray | sort -Unique							# Get unique value
            $i					= 1
			foreach ($varName in $varArray)
			{
				$value 			= "Get-OVLogicalInterconnectGroup -name '{0}'" -f $varName
				$variable 		= "lig$i"
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix $variable -Value $value -indentlevel 1))

				$ligGroupMapping 	= $ligGroupMapping.replace($varName, $variable)
				$i++
			}
			 
			# 2- We build the hash table
			
			$ligGroupMapping	= '@{' + $ligGroupMapping.replace($sepChar,';') + '}'		# Hash Table
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'ligGroupMapping' -Value $ligGroupMapping -indentlevel 1))

			$ligMappingParam 	= ' -LogicalInterconnectGroupMapping $ligGroupMapping'

		}

		# OSdeployment
		$deploymentTypeParam 	= ''
		switch ($deploymentMode)
		{
			'External' 
				{
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'deployNetwork' -Value ('Get-OVNetwork -name "{0}"' -f $deploymentNetwork) -indentlevel 1))
					$deploymentTypeParam 	= ' -DeploymentNetworkType {0} -DeploymentNetwork $deployNetwork ' -f $deploymentMode
				}
			'Internal'
			{
					$deploymentTypeParam 	= ' -DeploymentNetworkType {0} ' -f $deploymentMode
			}
		}
		# Config Script


		# Ensure there is No space afer backtick
		newLine
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix $( 'New-OVEnclosureGroup	-Name $name -enclosureCount $enclosureCount {0}{1} `' -f $v4AddressPoolParam, $v6AddressPoolParam) -isVar $False -indentlevel 1))
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('{0}{1} `' 	-f $ligMappingParam, $powerModeParam) -isVar $False -indentlevel 6))
		if ($deploymentTypeParam)
		{
			[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('{0} `'	-f $deploymentTypeParam) -isVar $False -indentlevel 6))
		}
		newLine # to end the command
		
		# --- Scopes
		if ($scopes)
		{
			newLine
			[void]$scriptCode.Add( (Generate-CustomVarCode -Prefix 'object' -Value 'get-OVEnclosureGroup | where name -eq $name' -indentlevel 1))
			generate-scopeCode -scopes $scopes -indentlevel 1

		}


		endIfBlock

		# Skip creating because resource already exists
		elseBlock
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW {0} already exists.' -f $name ) -isVar $False -indentlevel 1 ))
		endElseBlock

		newLine
        
        
	}
	
	if ($List)
	{
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	}
	
	 # ---------- Generate script to file
	 writeToFile -code $ScriptCode -file $ps1Files

}


# ---------- Logical Enclosure 
Function Import-LogicalEnclosure([string]$sheetName, [string]$WorkBook, [string]$subdir)
{

	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	$ListPS1files	 		= [System.Collections.ArrayList]::new()
	
	foreach ( $le in $List)
	{

		$name          		= $le.name
		$enclosureSN      	= $le.enclosureSerialNumber
		$enclosureName   	= $le.enclosureName
		$enclosureGroup		= $le.enclosureGroup
		$fwBaseline 		= $le.fwBaseline
		$fwInstall     		= $le.fwInstall
		$scopes			    = $le.scopes

		# Create logicalEnclosure filename here per LE
		$filename 			= "$subdir\" + $name.Trim().Replace($Space, '') + '.ps1'
		[void]$ListPS1files.Add($filename)

		$scriptCode         = [System.Collections.ArrayList]::new()
		connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating logical enclosure {0} "' -f $name) -isVar $False ))
 
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'le' 						-Value ("get-OVLogicalEnclosure | where name -eq  '{0}' " -f $name ) ))		
		
		ifBlock			-condition 'if ($le -eq $Null)' 	
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('# -------------- Attributes for logical enclosure {0} ' -f $name) -isVar $False -indentlevel 1))
		

		# rename enclosure
		if ($enclosureName)
		{
			$SNArray 			= "@('" + $enclosureSN.Replace($SepChar,"'$Comma'") 	+ "')"
			$nameArray 			= "@('" + $enclosureName.Replace($SepChar,"'$Comma'")   + "')"

			newLine
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '# -- Renaming enclosures  ' -isVar $False -indentlevel 1))
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix serialNumbers  	-value $SNArray -indentlevel 1))
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix enclosureNames  	-value $nameArray -indentlevel 1))

			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'for ($i=0; $i -lt $serialNumbers.Count; $i++)' -isVar $False -indentlevel 1))
			startBlock -indentlevel 1
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'this_enclosure' -Value 'Get-OVEnclosure | where serialNumber -Match $serialNumbers[$i]'  -indentlevel 2))
			ifBlock -condition 'if ( ($this_enclosure) -and ($enclosureNames -notcontains ($this_enclosure.Name) ) ) ' -isVar $False -indentlevel 2
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'Set-OVEnclosure -inputObject $this_enclosure -Name $enclosureNames[$i]' -isVar $False -indentlevel 3))	
			endIfBlock -indentlevel 2
			endBlock -indentlevel 1
		}

		# --- Enclosure		
		$enclosure 				= $enclosureSN.Split($SepChar)[0]
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix enclosure -value ("Get-OVEnclosure | where serialNumber -match '{0}' " -f $enclosure) -indentlevel 1 ))
		$enclParam 				= ' -Enclosure $enclosure'

		# --- EnclosureGroup
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix enclosureGroup -value ("Get-OVEnclosureGroup -name '{0}' " -f $enclosureGroup) -indentlevel 1 ))
		$egParam 				= ' -EnclosureGroup $enclosureGroup'


		# fwBaseline
		$fwParam 				= $Null
		if ($fwBaseline)
		{
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix fwBaseline -value ("Get-OVBaseline -SPPname '{0}' " -f $fwBaseline) -indentlevel 1 ))
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix fwInstall -value ('${0}' -f $fwInstall) -indentlevel 1 ))
			$fwParam = ' -FirmwareBaseline $fwBaseline -ForceFirmwareBaseline $fwInstall'
		}

		newLine
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('New-OVLogicalEnclosure -Name "{0}" {1}{2}{3}' -f $name, $enclParam, $egParam, $fwParam ) -isVar $false -indent 1) )

		# --- Scopes
		if ($scopes)
		{
			newLine
			[void]$scriptCode.Add( (Generate-CustomVarCode -Prefix 'object' -Value 'get-OVLogicalEnclosure | where name -eq $name' -indentlevel 1))
			generate-scopeCode -scopes $scopes -indentlevel 1

		}



		endIfBlock
		# Skip creating because resource already exists
		elseBlock
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW {0} already exists.' -f $name ) -isVar $False -indentlevel 1 ))
		endElseBlock

		newLine
	
		# Add disconnect and close the file
		[void]$ScriptCode.Add('Disconnect-OVMgmt')


		# ---------- Generate script to file
		writeToFile -code $ScriptCode -file $filename
		
	}
	
	return $ListPS1Files
}
	



# ---------- Profile and Template are in one function 
Function Import-ProfileorTemplate([string]$sheetName, [string]$WorkBook, [string]$ps1Files, [Boolean]$isSpt, [string]$subdir )
{
    $scriptCode             	= [System.Collections.ArrayList]::new()

	$cSheet,$spSheet,$connSheet,$localStorageSheet,$SanStorageSheet,$iLOSheet 		= $sheetName.Split($SepChar)

	$spList 					= if ($spSheet)	 			{get-datafromSheet -sheetName $spSheet -workbook $WorkBook				} else {$null}
	$connList 					= if ($connSheet) 			{get-datafromSheet -sheetName $connSheet -workbook $WorkBook 			} else {$null}
	$localStorageList 			= if ($localStorageSheet)	{get-datafromSheet -sheetName $localStorageSheet -workbook $WorkBook 	} else {$null}	
	$sanStorageList 			= if ($sanStorageSheet) 	{get-datafromSheet -sheetName $sanStorageSheet -workbook $WorkBook		} else {$null}
	$iLOList 					= if ($iLOSheet) 			{get-datafromSheet -sheetName $iLOSheet -workbook $WorkBook				} else {$null}


	$isSP 						= -not $isSpt

	if ($null -ne $localStorageList)
	{
		# Define class for sasJBOD
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'class sasJBOD' -isVar $False -indentlevel 0 ))
		startBlock
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '[int]$id'                    -isVar $False -indentlevel 1 ))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '[string]$deviceSlot'         -isVar $False -indentlevel 1 ))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '[string]$name'               -isVar $False -indentlevel 1 ))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '[string]$description'        -isVar $False -indentlevel 1 ))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '[int]$numPhysicalDrives'     -isVar $False -indentlevel 1 ))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '[string]$driveMinSizeGB'     -isVar $False -indentlevel 1 ))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '[string]$driveMaxSizeGB'     -isVar $False -indentlevel 1 ))	
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '[string]$driveTechnology'    -isVar $False -indentlevel 1 ))	
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '[boolean]$eraseData'         -isVar $False -indentlevel 1 ))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '[boolean]$persistent'        -isVar $False -indentlevel 1 ))
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '[string]$sasLogicalJBODUri'  -isVar $False -indentlevel 1 ))
		endBlock
	}


	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode
	if ($isSP)
	{
		$count		 				= 1
		$MAXPROFILE					= 10
		$i							= 0
		$ListPS1Files				= [System.Collections.ArrayList]::new()
		# Build a first set
		$spFilename 				= "$subdir\profileGroup$count" + ".PS1"
	}



	foreach ( $prof in $spList)
	{

		if ($isSP) 				{$i++}   # counter for number of profiles to be in a ps1 file
		$spName 				= $prof.name
		$description    		= $prof.Description 
		$spDescription      	= $prof.serverprofileDescription 
		$template	  			= $prof.serverProfileTemplate
		$serverHardware			= $prof.serverHardware
		$sht 					= $prof.serverHardwareType
		$eg 					= $prof.enclosureGroupName
		$affinity 				= $prof.affinity

		$manageFirmware			= [Boolean]($prof.manageFirmware)
		$fwBaseline				= $prof.firmwareBaselineName
		$fwInstallType 			= $prof.firmwareInstallType
		$fwForceInstall			= $prof.forceInstallFirmware
		$fwActivation 			= $prof.firmwareActivationType
		$fwSchedule 			= $prof.firmwareSchedule

		$bm                 	= [Boolean]($prof.manageBootMode)
		$bmMode 				= $prof.mode
		$pxeBootPolicy			= $prof.pxeBootPolicy
		$secureBoot 			= $prof.secureBoot
		$bo                 	= [Boolean]($prof.manageBootOrder)
		$order 					= $prof.order

		$bios 					= [Boolean]($prof.manageBios)
		$biosSettings	 		= $prof.overriddenSettings

		$manageConnections		= [Boolean]($prof.manageConnections)
		$manageSANStorage 		= [Boolean]($prof.manageSANStorage)

		$manageIlo 				= [Boolean]($prof.manageIlo)

		if ($isSpt)
		{
			$fwConsistency 		= if ($prof.firmwareConsistencyChecking)  	{$consistencyCheckingEnum.Item($prof.firmwareConsistencyChecking) } else {'None'}
			$bmConsistency 		= if ($prof.bootModeConsistencyChecking) 	{$consistencyCheckingEnum.Item($prof.bootModeConsistencyChecking) } else {'None'}
			$boConsistency 		= if ($prof.bootOrderConsistencyChecking)	{$consistencyCheckingEnum.Item($prof.bootOrderConsistencyChecking) } else {'None'}
			$biosConsistency 	= if ($prof.biosConsistencyChecking) 		{$consistencyCheckingEnum.Item($prof.biosConsistencyChecking) } else {'None'}
			$lsConsistency 		= if ($prof.localStorageConsistencyChecking)	{$consistencyCheckingEnum.Item($prof.localStorageConsistencyChecking) } else {'None'}
			$sanConsistency 	= if ($prof.sanStorageConsistencyChecking) 	{$consistencyCheckingEnum.Item($prof.sanStorageConsistencyChecking) } else {'None'}
			$connConsistency 	= if ($prof.connectionConsistencyChecking) 	{$consistencyCheckingEnum.Item($prof.connectionConsistencyChecking) } else {'None'}
			$iloConsistency 	= if ($prof.iloConsistencyChecking)			{$consistencyCheckingEnum.Item($prof.iloConsistencyChecking) } else {'None'}
		}



		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating profile {0} "' -f $spName) -isVar $False ))
 
		$getCmd		= $newCmd =  $saveCmd = $null
		if ($isSpt)
		{
			$getCmd 			= 'get-OVServerProfileTemplate'
			$newCmd 			= 'new-OVServerProfileTemplate'
			$saveCmd 			= 'save-OVServerProfileTemplate'
		}
		else 
		{
			$getCmd 			= 'get-OVServerProfile'
			$newCmd 			= 'new-OVServerProfile'
			$saveCmd 			= 'save-OVServerProfile'
		}

		$value 					= $getCmd + " | where name -eq  '{0}' " -f $spName
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'profile' 						-Value $value ))		
		
		ifBlock			-condition 'if ($Null -eq $profile )' 
		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('# -------------- Attributes for profile {0} ' -f $name) -isVar $False -indentlevel 1))

		[void]$scriptCode.Add((Generate-CustomVarCode -Prefix name -value ("'{0}'" -f $spName) -indentlevel 1 ))
		
		$descParam 					=  $null
		if ($description) 
		{
			$descParam 				=  ' -description $description'
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix description -value ("'{0}'" -f $description) -indentlevel 1 ))
		}

		# ############### Server Profile Region
		$spDescParam 				=  $null
		if ( $isSpt -and $spDescription)
		{
			$spdescParam 			= ' -ServerProfileDescription $spDescription'
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix spDescription -value ("'{0}'" -f $spDescription) -indentlevel 1 ))
		}

		# --- server profile template
		# Used when creating profile from template
		$spTemplateParam 	= $null
		if ($template)
		{
			$spTemplateParam = ' -ServerProfileTemplate $spTemplate'
			$value 			 = "Get-OVServerProfileTemplate -name '{0}'" -f $template
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix spTemplate -value $value -indentlevel 1 ))
		}
		else # Standalone profile
		{
			# -- server hardware type	
			$shtParam 			= ' -ServerHardwareType $sht'
			$value 				= "Get-OVserverHardwareType -name '{0}'" -f $sht

			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix sht -value $value -indentlevel 1 ))

			# -- enclosure group
			$egParam 			= ' -EnclosureGroup $eg'
			$value 				= "Get-OVEnclosureGroup -name '{0}'" -f $eg
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix eg -value $value -indentlevel 1 ))

			# --- affinity
			$affinityParam 		= ' -affinity $affinity'
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix affinity -value ("'{0}'" -f $affinity) -indentlevel 1 ))

		}

		# ---- server hardware
		$hwParam 			= $null
		if ($serverHardware)
		{
			$hwParam 		= ' -Server $server'
			$value 			 = "Get-OVServer -name '{0}'" -f $serverHardware
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix server -value $value -indentlevel 1 ))
			# -- Add code to power off server
			$value 			= 'Stop-OVServer -inputObject $server -force -Confirm:$False| Wait-OVTaskComplete	'
			[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix status -value $value  -indentlevel 1))
				
		}

	if ($Null -eq $template) 			# Standalone SP or SPT
	{
			# ############### OS Deployment Region



			# ############### Firmware Region
			$fwParam 				= $null
			if ($manageFirmware)
			{
				if ($null -eq $fwActivation)
				{
					$fwActivation 			= 'NotScheduled'
				}

				if ($null -eq $fwInstallType)
				{
					$fwInstallType		= 'FirmwareAndOSDrivers'
				}
				newLine
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '# -------------- Firmware Baseline section ' -isVar $False -indentlevel 1))
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'manageFirmware' -Value ('${0}' -f $manageFirmware) -indentlevel 1))
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'sppName' -Value ('"{0}"' -f $fwBaseline) -indentlevel 1))
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'fwBaseline' -Value 'Get-OVbaseline -SPPname $sppName' -indentlevel 1))
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'fwInstallType' -Value ('"{0}"' -f $fwInstallType) -indentlevel 1))
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'fwForceInstall' -Value ('${0}' -f $fwforceInstall)-indentlevel 1))
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'fwActivation' -Value ('"{0}"' -f $fwActivation) -indentlevel 1))
				
				$fwParam = ' -firmware -Baseline $fwBaseline -FirmwareInstallMode $fwInstallType -ForceInstallFirmware:$fwForceInstall -FirmwareActivationMode $fwActivation '

				$fwConsistencyParam	= $fwScheduleParam = $null
				if ($isSpt)
				{
					if ($null -eq $fwConsistency)
					{
						$fwConsistency 	= 'None'
					}
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'fwConsistency' -Value ('"{0}"' -f $fwConsistency) -indentlevel 1))
					$fwConsistencyParam	= ' -FirmwareConsistencyChecking $fwConsistency'
				}
				else   # SP specific here
				{
					if ($fwSchedule)
					{
						[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'fwSchedule' -Value ('[DateTime]{0}' -f $fwSchedule) -indentlevel 1))
						$fwScheduleParam = ' -FirmwareActivateDateTime $fwSchedule'
					}
				}

				$fwParam += $fwConsistencyParam	+ $fwScheduleParam 
			}


			# ############### Connections
			newLine
			[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '# -------------- Connections section ' -isVar $False -indentlevel 1))

			$connectionsParam 			= $null
			if ($manageConnections)
			{
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'manageConnections' -Value ('${0}' -f $manageConnections) -indentlevel 1))
				$connectionsParam  	 	= ' -manageConnections $manageConnections'
			
				if ($isSpt)
				{
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'connectionsConsistency' -Value ('"{0}"' -f $connConsistency) -indentlevel 1))
					$connectionsParam	+= ' -ConnectionsConsistencyChecking $connectionsConsistency'
				}
			}

				$connectionArray       	= [System.Collections.ArrayList]::new()
				$connectionList 		= $connList | where ProfileName -eq $spName
				$index					= 1

				foreach ($conn in $connectionList)
				{
					$name 				= $conn.name
					$id			        = $conn.id
					$functionType 		= $conn.functionType
					$network 			= $conn.network
					$portId       		= $conn.portId
					$requestedMbps		= $conn.requestedMbps
					$requestedVFs 		= $Conn.requestedVFs

					$lagName	 		= $conn.lagName


					$bootable			= $conn.boot
					$priority       	= $conn.priority
					$bootVolumeSource 	= $conn.bootVolumeSource


					if ($isSp)
					{
						$userDefinedParam 	= ''
						$userDefined 	= $conn.userDefined
						if ($userDefined)
						{
							$macType 	= $Conn.macType
							$mac	 	= $Conn.mac
							$wwpnType	= $Conn.wwpnType
							$wwpn		= $Conn.wwpn
							$wwnn		= $Conn.wwnn
							$_macParam 	= if ($mac)		{ ' -mac {0} '  -f $mac} 	else {''}
							$_wwnnParam = if ($wwnn)	{ ' -wwnn {0} ' -f $wwnn}	else {''}
							$_wwpnParam = if ($wwpn)	{ ' -wwpn {0} ' -f $wwnn}	else {''}

							if ($_macParam  -or $_wwnnParam -or $_wwpnParam)
							{
								$userDefinedParam = ' -UserDefined:$True ' + $_macParam + $_wwpnParam + $_wwnnParam
							}
						}

					}


					$value 			= "Get-OVnetwork | where name -eq '{0}' " -f $network		
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'network' 						-Value $value -indentlevel 1))		
					ifBlock -condition 'if ($null -eq $network)'  -indentlevel 1
					$value 			= "Get-OVnetworkSet | where name -eq '{0}' " -f $network		
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'network' 						-Value $value -indentlevel 2))
					endIfBlock -indentLevel 1



					if ($bootable)
					{
						$bootParam 	= ' -bootable:${0} -priority {1} -bootVolumeSource {2} ' -f $bootable,$priority,$bootVolumeSource
					}

					# TBD FibreChannel Bfs

					$value 				+= ' -network $network'
					$value 				+= $bootParam

					$nameParam 			= if ($name) {' -name "{0}" ' -f $name} else {''}

					# -- code
					$_connection		= '$' + "conn$index"
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ( '{0}   = New-OVServerProfileConnection {1} -ConnectionID {2} `' 	-f $_connection,$nameParam, $id)  		-isVar $False  -indentlevel 1))
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ( '  -PortId "{0}" -RequestedBW {1} `' 								-f $portId , $requestedMbps) 		-isVar $False  -indentlevel 11))
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ( '  -network $network {0} {1} ' 										-f $bootParam, $userDefinedParam)	-isVar $False  -indentlevel 11))
					
					newLine
					[void]$connectionArray.Add($_connection)
					$index++

				}

				if ($connectionArray)
				{
					$value 					= '@(' + ($connectionArray -join $comma) + ')'
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'connectionList' -Value $value -indentlevel 1))
					$connectionsParam 		= ' -Connections $connectionList '
					$connectionArray       	= [System.Collections.ArrayList]::new()
				}
			


			# ############### Local Storage

			$localStorageParam				= $null
			$lsList 						= $localStorageList  | where ProfileName -eq $spName

			if ($Null -ne $lsList)
			{
				newLine
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '# -------------- Local Storage section ' -isVar $False -indentlevel 1))
				if ($isSpt)
				{
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'lsConsistencyChecking' -Value ('"{0}"' -f $lsConsistency) -indentlevel 1))
					$localStorageParam		= ' -LocalStorageConsistencyChecking $lsConsistencyChecking'
				}
				newLine
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '# --- Search for SAS-Logical-Interconnect for SASlogicalJBOD ' -isVar $False -indentlevel 1))
				
				# Find SAS-Logical-INTERCONNECT from logical enclosure
				# Note: $eg is defined earlier in the generated script

				# Step 1 - find SasLIG from EnclosureGroup
				$value 				= 'Search-OVAssociations ENCLOSURE_GROUP_TO_LOGICAL_INTERCONNECT_GROUP -parent $eg' 
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ligAssociation -value $value -indentlevel 1 ))

				$value 				= '($ligAssociation | where ChildUri -like "*sas*").ChildUri'
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix sasLigUri -value $value -indentlevel 1 ))

				$value 				= 'if ($sasLigUri) { Send-OVRequest -uri $sasLigUri } else {$Null}'
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix sasLig -value $value -indentlevel 1 ))

				# Step 2 - find Sas Interconnect from SAS Lig
				$value 				= '(Search-OVAssociations LOGICAL_INTERCONNECT_GROUP_TO_LOGICAL_INTERCONNECT -Parent $sasLig).childUri[0]' 
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix sasUri -value $value -indentlevel 1 ))

				$value 				= 'if ($sasUri) { Send-OVRequest -uri $sasUri } else {$Null}'
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix sasLI -value $value -indentlevel 1 ))

				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix lsJBOD -value '[System.Collections.ArrayList]::new()' -indentlevel 1 ))
			
				$controllerArray		= [System.Collections.ArrayList]::new()
				$logicaldiskArray 		= [System.Collections.ArrayList]::new()
				$storageVolumeArray 	= [System.Collections.ArrayList]::new()
				

				$diskIndex 				= 1	
				$volumeIndex 			= 1
				$contIndex 				= 1
				$jbodIndex				= 1
				$iloUserIndex 			= 1
				$ilodirectoryGroupIndex = 1



				foreach ($ls in $lsList)
				{
					$deviceSlot				= $ls.deviceSlot
					$mode 					= $ls.mode
					$initialize				= [Boolean]($ls.initialize)
					$writeCache				= if ($null -eq $ls.driveWriteCache) {'Unmanaged'} else {$ls.driveWriteCache} 
					
					$logicalDrives			= $ls.logicalDrives			#[]
					$numPhysicalDrives		= [string]$ls.numPhysicalDrives		#[]
					$raidLevel 				= $ls.raidLevel				#[]
					$bootable				= $ls.bootable				#[]
					$accelerator			= $ls.accelerator			#[]
					$driveTechnology 		= $ls.driveTechnology		#[]
					#SasLogicalJBOD
					$driveID 				= $ls.id
					$driveDescription		= if ($ls.description) { $ls.description} else {''}
					$driveMinSize			= if ($ls.driveMinSize) {$ls.driveMinSize} else {0}
					$driveMaxSize			= if ($ls.driveMaxSize) {$ls.driveMaxSize} else {0}
					$eraseData 				= $ls.eraseData
					$persistent 			= $ls.persistent



					# Internal Disks only
					if ($mode)
					{
						# Check logical drives
						if ($logicalDrives)
						{
							$logicalDrivesArray	= $logicalDrives.Split('|')
							$physDrivesArray 	= $numPhysicalDrives.Split('|')
							$raidLevelArray 	= $raidLevel.Split('|')
							$bootableArray 		= $bootable.Split('|')
							$acceleratorArray 	= $accelerator.Split('|')
							$driveTypeArray		= $driveTechnology.Split('|') 
								
							# Get logical drives first
							foreach ($_ld in $logicalDrivesArray)
							{	

								$ldParam = $ldSizeParam = $ldLocParam = $null

								newLine
								$prefix 		= '# --- Attributes for Logical Disk {0} ({1})' -f $_ld, $deviceSlot
								[void]$scriptCode.Add((Generate-CustomVarCode -Prefix $prefix	-isVar $False 	-indentlevel 1))
								
								$value 			= "'{0}'" -f $_ld
								[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'ldName'			-value $value 	-indentlevel 1))
								$ldParam 		+= ' -Name $ldName'
								
								$this_index 	= [array]::IndexOf($logicalDrivesArray, $_ld )


								$value 			= '${0}' -f $bootableArray[$this_index] 
								[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'bootable'		-value $value 	-indentlevel 1))
								$ldParam 		+= ' -Bootable $bootable'
								
								$value 			= "'{0}'" -f $raidLevelArray[$this_index]	
								[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'raidLevel'		-value $value 	-indentlevel 1))
								$ldParam 		+= ' -RAID $raidLevel'

								$_driveType 	= $driveTypeArray[$this_index]
								if ($_driveType)
								{
									$value 			= "'{0}'" -f $_driveType	
									[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'driveType'		-value $value 	-indentlevel 1))	
									$ldParam 		+= ' -DriveType $driveType' 
								}

								$value 			= "{0}" -f $physDrivesArray[$this_index]
								[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'numberofDrives'	-value $value 	-indentlevel 1))	
								$ldParam 		+= ' -NumberofDrives $numberofDrives '


								$value 			= "'{0}'" -f $acceleratorArray[$this_index]	
								[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'accelerator'		-value $value 	-indentlevel 1))
								$ldParam 		+=' -Accelerator $accelerator'



								# Make sure that there is no space after backstick (`)
								$logicalDisk 	= '{0}' -f "LogicalDisk$diskIndex"
								[void]$scriptCode.Add((Generate-CustomVarCode -Prefix $logicalDisk  -value ('New-OVServerProfileLogicalDisk {0} ' -f $ldParam)	-indentlevel 1))

								[void]$logicaldiskArray.Add('${0}' -f $logicalDisk)
								$diskIndex++
							}
						}
					
				
						
						# --- Generate array of logical disk for this controller
						if ($logicaldiskArray)
						{
							$logicalDisks 		= '@(' + ($logicaldiskArray -join $comma) + ')'
							[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'logicalDisks'  -value $logicalDisks -indentlevel 1))
							$logicalDiskParam 	= ' -LogicalDisk $LogicalDisks '
						}
						else 
						{
							$logicalDiskParam 	= ''
						}
			

						$controllerParam  		= ' -ControllerID "{0}" -Mode "{1}" -Initialize:${2} -WriteCache "{3}" {4}' -f $deviceSlot, $mode, $initialize, $writeCache, $logicalDiskParam 
						$controller				= '{0}' -f "controller$contIndex"  

						# ---- Generate new Disk Controller
						[void]$scriptCode.Add((Generate-CustomVarCode -Prefix $controller  -value ('New-OVServerProfileLogicalDiskController {0}' -f $controllerParam) -indentlevel 1))
						newLine

						[void]$controllerArray.Add('${0}' -f $controller)
						$contIndex++
						$logicaldiskArray 		= [System.Collections.ArrayList]::new()
					}
					else # SasLogicalJBOD
					{
						$jbodDisk 			= '$' + "jbod$jbodIndex"
						

						ifBlock 	-condition 'if ($sasLI)			# If SAS logical Interconnect exists' -indentlevel 1
							$value 					= 'Get-OVAvailableDriveType -InputObject $sasLI | where { $_.Capacity -eq $MaxDriveSize -and $_.Type -eq $DriveTechnology -and $_.NumberAvailable -eq $numPhysicalDrives}'
							[void]$scriptCode.Add((Generate-CustomVarCode -Prefix availableDrives  -value $value -indentlevel 2))
							ifBlock	-condition 'if ($availableDrives)		# if there are such drives in SAS' -indentlevel 2
								$JBODParam 			= ' -name {0} -driveType {1} -MinDriveSize {2} -MaxDriveSize {3} -EraseDataOneDelete ${4} ' -f $logicalDrives, $driveTechnology, $driveMinSize, $driveMaxSize, $eraseData
								
								if ($isSP) ##### For server profile only. Check for existing JBOD and attach to server profile.
								{
									$value 				= 'Get-OVLogicalJBOD -name "{0}" ' -f $logicalDrives
								}
								else
								{
									$value 				= 'new-OVLogicalJBOD -InputObject $sasLI {0} ' -f $JBODParam
								}
								$jbodDisk 			= '$' + "jbod$jbodIndex"
								[void]$scriptCode.Add((Generate-CustomVarCode -Prefix $jbodDisk  -value $value -isVar $False -indentlevel 3))
								$jbodIndex++
								ifBlock -condition "if ($jbodDisk)"  -indentlevel 3
									[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '_jbod' 					-value 'new-object -type sasJBOD' 	-indentlevel 4))
									[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '_jbod.id'                -value $driveID                     -indentlevel 4 ))
									[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '_jbod.deviceSlot'        -value $deviceSlot                  -indentlevel 4 ))
									[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '_jbod.name'              -value $logicalDrives               -indentlevel 4 ))
									if ($driveDescription)
									{
										[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '_jbod.description'   -value $driveDescription            -indentlevel 4 ))
									}
									[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '_jbod.numPhysicalDrives' -value $numPhysicalDrives           -indentlevel 4 ))
									[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '_jbod.driveMinSizeGB'    -value $driveMinSize                -indentlevel 4 ))
									[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '_jbod.driveMaxSizeGB'    -value $driveMaxSize                -indentlevel 4 ))	
									[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '_jbod.driveTechnology'   -value $driveTechnology             -indentlevel 4 ))	
									[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '_jbod.eraseData'         -value $eraseData                   -indentlevel 4 ))
									[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '_jbod.persistent'        -value $persistent                  -indentlevel 4 ))
									[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '_jbod.sasLogicalJBODUri' -value ('{0}.uri' -f $jbodDisk)     -indentlevel 4 ))

									[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '[void]$lsJBOD.Add($_jbod)'  -isVar $False 					-indentlevel 4 ))

								endIfBlock -indentlevel 3							
							endIfBlock -indentlevel 2
						endIfBlock -indentlevel 1

					}

				}

				if ($controllerArray)
				{
					# ----- Generate params for profiles
					$controllers 	= '@(' + ($controllerArray -join $comma) + ')'
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'controllers'  -value $controllers  -indentlevel 1))
					$localStorageParam 	= ' -StorageController $controllers'
					$controllerArray		= [System.Collections.ArrayList]::new()
					$logicaldiskArray 		= [System.Collections.ArrayList]::new()
				}


			}

			# ############### SAN Storage
			$StorageVolumeParam				= $null


			$volList 						= $sanStorageList  | where ProfileName -eq $spName
			if ($volList)
			{
				newLine
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '# -------------- SAN Storage section ' -isVar $False -indentlevel 1))
				if ($isSpt)
				{
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'sanConsistencyChecking' -Value ('"{0}"' -f $sanConsistency) -indentlevel 1))
					$StorageVolumeConsistencyParam	= ' -sanStorageConsistencyChecking $sanConsistencyChecking'
				}
				newLine

				foreach ($_vol in $volList)
				{
					$_volName 					= $_vol.volumeName
					$_volUri 					= '(Get-OVStorageVolume -name "{0}" ).uri' -f $_volName

					$_lunType 					= $_vol.volumeLunType
					$_lun						= $_vol.volumeLUN

					$_sts						= $_vol.volumeStorageSystem
					$_stsUri 					= '(Get-OVStorageSystem -name "{0}" ).uri' -f $_sts

					$_stPath 					= $_vol.volumeStoragePaths
					if ($_stPath)
					{	
						$_stPath 				= '@( ' + $_stPath.replace($SepChar, $COMMA)  + ' )' 		# Build array 
					}

					$_volVariable 				= '$volume{0}'  -f $volumeIndex++
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix "$_volVariable	=	@{" 	-isVar $False 	-indentlevel 1))
					if ($_lunType -eq 'Manual')
					{
						[void]$scriptCode.Add((Generate-CustomVarCode -Prefix "id						= $_lun $SepHash"  		-isVar $False 	-indentlevel 6))
					}
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix "lunType					= '$_lunType' $SepHash"	-isVar $False 	-indentlevel 6))
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix "volumeUri				= $_volUri $SepHash" 	-isVar $False 	-indentlevel 6))
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix "volumeStorageSystemUri	= $_stsUri $SepHash" 	-isVar $False 	-indentlevel 6))			
					if ($_stPath)
					{
						[void]$scriptCode.Add((Generate-CustomVarCode -Prefix "storagePaths			= $_stPath " 			-isVar $False 	-indentlevel 6))
					}
					[void]$scriptCode.Add((Generate-CustomVarCode 	-Prefix '}'												-isVar $False 	-indentlevel 6))

					[void]$storageVolumeArray.add($_volVariable)

				}
				if ($storageVolumeArray )
				{
					$value 						= "@( " + ($storageVolumeArray -join $Comma) + " )"
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix storageVolume  	-value $value	-indentlevel 1))
					$StorageVolumeParam 		= ' -sanStorage -StorageVolume $storageVolume '
					
					if ($isSpt)
					{
						$StorageVolumeParam		+=  $StorageVolumeConsistencyParam
					}
					$storageVolumeArray 		= [System.Collections.ArrayList]::new()
				}
			}
			



			# ############### Boot Settings : BootMode / BootOrder
			$bmParam 		= $bootParam = $null
			if ($bm )
			{
				newLine
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '# -------------- Boot mode section ' -isVar $False -indentlevel 1))
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'bootMode' 		-Value ('"{0}"' -f $bmMode) -indentlevel 1))
				

				$bmParam 		= ' -bootMode $bootMode '
				if ($pxeBootPolicy)
				{
					$bmParam 	+=  ' -PxeBootPolicy {0}' -f  $pxeBootPolicy
				}

				if ($bootMode -match 'UEFI Optimized*')
				{
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'secureBoot' 		-Value ('"{0}"' -f $secureBoot) -indentlevel 1))
					$bmParam 	+= ' -SecureBoot $secureBoot'
				}

				if ($order)
				{
					$bootOrder 	 = "@('" + $order.Replace($SepChar,"'$Comma'") + "')" 
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'bootOrder' 		-Value ('{0}' -f $bootOrder) -indentlevel 1)) 
					$boParam     = ' -ManageBoot ${0} -BootOrder $bootOrder' -f $bo
				}

				$bmConsistencyParam	= $null
				$boConsistencyParam = $null

				if ($isSpt)
				{
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'bmConsistency' -Value ('"{0}"' -f $bmConsistency) -indentlevel 1))
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'boConsistency' -Value ('"{0}"' -f $boConsistency) -indentlevel 1))
					
					$bmConsistencyParam	= ' -BootModeConsistencyChecking $bmConsistency'
					$boConsistencyParam	= ' -BootOrderConsistencyChecking $boConsistency'

					$bmParam 			+= $bmConsistencyParam	
					$boParam 			+= $boConsistencyParam
				}
			}

			# ############### BIOS Settings
			$biosParam 		= $null
			if ($bios)
			{

				# Build Array of HashTables
				$settingsArr 	= $biosSettings.Split($SepChar)

				newLine
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '# -------------- BIOS section ' -isVar $False -indentlevel 1))
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'biosSettings' 	-Value '@(' -indentlevel 1))

				foreach ($setting in $settingsArr)
				{
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ("{0}$COMMA" -f $setting)	-isVar $false  -indentlevel 2))
				}
				$scriptCode[-1] = $scriptCode[-1] -replace $COMMA , ''
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ')'	-isVar $false  -indentlevel 2))



				$biosParam 		= ' -Bios -BiosSettings $biosSettings'

				if ($isSpt)
				{
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'biosConsistency' -Value ('"{0}"' -f $biosConsistency) -indentlevel 1))				
					$biosParam 	+= ' -BiosConsistencyChecking $biosConsistency'	
				}
			}


			# ############### iLO Settings
			$iloParam 					= '' 
			if ($manageIlo)
			{

				$iloLocalAccounts		= [System.Collections.ArrayList]::new()
				$iloDirectoryGroups		= [System.Collections.ArrayList]::new()
				$iloParam = $iloAdminParam	= $iloLocalAccountsParam = $iloHostnameParam = ''

				$dirIndex 				= 1

				$iLOsettingList 		= $iLOList | where ProfileName -match $spName

				newLine
				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix '# -------------- iLO section ' -isVar $False -indentlevel 1))

				[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'manageIlo' -Value ('${0}' -f $manageIlo) -indentlevel 1))
				$iloParam  	 			= ' -ManageIloSettings $manageIlo'
				if ($isSpt)
				{
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix 'iloConsistency' -Value ('"{0}"' -f $iloConsistency) -indentlevel 1))				
					$iloParam 			+= ' -IloSettingsConsistencyChecking $iloConsistency'	
				}


				foreach ($s in $iLOsettingList)
				{

					switch ($s.settingType)
					{
						'AdministratorAccount'
							{
								$iloAdminParam 			 = ' -ManageLocalAdministratorAccount '
								$iloAdminParam 			+= ' -RemoveLocalAdministratorAccount ${0} ' -f $s.deleteAdministratorAccount 
								if ($s.adminPassword)
								{
									$iloAdminParam 		+= ' -LocalAdministratorPassword $({0} | ConvertTo-SecureString -AsPlainText -Force )' -f $s.adminPassword
								}
							}					
						
						'LocalAccounts'
							{
								if ($s.username)
								{
									$_privParamArr		= [System.Collections.ArrayList]::new()

									$iloNameParam		= " -Username '{0}' "	 													-f $s.userName	
									$iloDisplayParam 	= if ($s.displayName)		{ " -DisplayName '{0}' "	 					-f $s.displayName	 }	else {''}
									$iloPasswordParam 	= ' -Password $("{0}" | ConvertTo-SecureString -AsPlainText -Force ) '		-f $s.userPassword
									
									$privileges 		= $s.userPrivileges
									if ($privileges)
									{
										$privArr 		= $privileges.Split($SepChar) 
										foreach ($priv in $privArr)
										{
											$_privParamArr	+= $iLOPrivilgeParamEnum.Item($priv)
										}
									}

								
									$iloUser 			= '$iloUser{0}' -f $iloUserIndex++
									$value 				= 'new-OVIloLocalUserAccount	{0}{1} `' 		-f $iloNameParam, $iloDisplayParam 
									[void]$scriptCode.Add((Generate-CustomVarCode -Prefix $iloUser 	-value $value 			-isVar $False	-indentlevel 1)) 
									[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0} `'		-f $iloPasswordParam)	-isVar $False	-indentlevel 18))
									foreach ($_param in $_privParamArr)
									{
										[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0} `'	-f $_param)		-isVar $False	-indentlevel 18))
									}
									
									newLine
									[void]$iloLocalAccounts.add($iloUser)
								
								}
								
								
							}
						
						'Directory'
							{
								$_dirParamArr					= [System.Collections.ArrayList]::new()

								$_p 							= ' -ManageDirectoryConfiguration '
								[void]$_dirParamArr.Add($_p)

								$dirAuth 						= $s.directoryAuthentication
								if ($dirAuth -ne 'Disabled')
								{
									$_p 						= ' -LdapSchema {0} -GenericLDAP ${1} ' -f $dirAuth, $s.directoryGenericLDAP
									[void]$_dirParamArr.Add($_p)
									$_p 						= ' -LOMObjectDistinguishedName "{0}" ' -f $s.iloObjectDistinguishedName
									[void]$_dirParamArr.Add($_p)
									
									if ($s.directoryPassword)
									{
										$_p 					= ' -IloObjectPassword $("{0}" | ConvertTo-SecureString -AsPlainText -Force) ' -f $s.directoryPassword
										[void]$_dirParamArr.Add($_p)
									}
									$_p 						= ' -DirectoryServerAddress {0} -DirectoryServerPort {1} ' -f $s.directoryServerAddress, $s.directoryServerPort
									[void]$_dirParamArr.Add($_p)

									$userContext 				= $s.directoryUserContext
									if ($userContext)
									{
										$userContext			= "'" + $userContext.replace($sepChar, "' , '") + "'" 
										$_p						= ' -DirectoryUserContext {0} ' -f $userContext
										[void]$_dirParamArr.Add($_p)
									}
								}

								$kerberosAuth 					= 	[Boolean]($s.kerberosAuthentication)
								if ($kerberosAuth)
								{
									$_p 						= ' -EnableKerberosAuthentication ${0} ' -f $kerberosAuth 
									[void]$_dirParamArr.Add($_p)
									$_p 						= ' -KerberosRealm {0} -KerberosKDCServerAddress {1} -KerberosKDCServerPort {2}   ' -f $s.kerberosRealm, $s.kerberosKDCServerAddress, $s.kerberosKDCServerPort
									[void]$_dirParamArr.Add($_p)
								}													
								
									#TBD - KerberosKeyTab



							}
						
						'DirectoryGroups'
							{
								if ($s.groupDN)
								{
									$_privParamArr				= [System.Collections.ArrayList]::new()
									$iloDirGroupParam1 			= ' -GroupDN "{0}" ' -f  $s.groupDN
									$iloDirGroupParam1 			+= if ($s.groupSID) { '-GroupSID {0}' -f $s.groupSID} else {''}
									$privileges 				= $s.groupPrivileges
									if ($privileges)
									{
										$privArr 				= $privileges.Split($SepChar) 							
										foreach ($priv in $privArr)
										{
											$_privParamArr			+= $iLOPrivilgeParamEnum.Item($priv)
										}
									}

									$iloDirectoryGroup 			= '$iloDG{0}' -f $ilodirectoryGroupIndex++
									[void]$scriptCode.Add((Generate-CustomVarCode -Prefix $iloDirectoryGroup 	-value ('New-OVIloDirectoryGroup{0} `' -f $iloDirGroupParam1) -isVar $False -indentlevel 1))
									foreach ($_param in $_privParamArr)
									{
										[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0} `'	-f $_param)		-isVar $False	-indentlevel 18))
									}	
									newLine
									[void]$iloDirectoryGroups.add($iloDirectoryGroup)
									
								}

							}	
							
						'HostName'
							{
								$iloHostnameParam 		= if ($s.hostName ) 	{' -ManageIloHostname -IloHostname "{0}" ' -f $s.hostName } else {''}
															
							}
						
						'KeyManager'
							{
								$_kmParamArr					= [System.Collections.ArrayList]::new()

								if ($s.primaryServerAddress)
								{
									[void]$_kmParamArr.Add(' -ManageKeyManager  ') 

									$_p 			= ' -PrimaryKeyServerAddress {0} -PrimaryKeyServerPort {1} ' 					-f $s.primaryServerAddress, $s.primaryServerPort
									[void]$_kmParamArr.Add($_p)
								}
								if ($s.secondaryServerAddress)
								{
									$_p				= ' -SecondaryKeyServerAddress {0} -SecondaryKeyServerPort {1} ' 				-f $s.secondaryServerAddress, $s.secondaryServerPort
									[void]$_kmParamArr.Add($_p)
								}
								if ($s.certificateName)
								{
									$_p				= ' -KeymanagerLocalCertificateName {0} '										-f $s.certificateName
									[void]$_kmParamArr.Add($_p)
								}

								$_p 				= ' -RedundancyRequired ${0} -KeymanagerGroupName {1}' 							-f $s.redundancyRequired, $s.groupName
								[void]$_kmParamArr.Add($_p)

								$securePassword 	= '$("{0}" | ConvertTo-SecureString -AsPlainText -Force )'						-f $s.keyManagerpassword
								$_p					= ' -KeymanagerLoginName {0} -KeymanagerPassword {1} '							-f $s.loginName, $securePassword
								[void]$_kmParamArr.Add($_p)

							}
						
					}


				}

				
				if ($iloLocalAccounts)
				{
					$iloLocalAccountsParam 		= ' -ManageLocalAccounts -LocalAccounts $iloLocalAccounts '
					[void]$scriptCode.Add((Generate-CustomVarCode 	-Prefix 'iloLocalAccounts' -value ('@({0})' -f ($iloLocalAccounts -join $COMMA))	-indentlevel 1))
				}
				
				if ($iloDirectoryGroups)
				{
					$iloDirectoryGroupsParam 	= ' -ManageDirectoryGroups -DirectoryGroups $iloDirectoryGroups '
					[void]$scriptCode.Add((Generate-CustomVarCode 	-Prefix 'iloDirectoryGroups' -value ('@({0})' -f ($iloDirectoryGroups -join $COMMA))	-indentlevel 1))
				}

				[void]$scriptCode.Add((Generate-CustomVarCode 	-Prefix 'iloPolicy' -value ('new-OVServerProfileIloPolicy	{0} `' -f $iloAdminParam)	-indentlevel 1))

				
				if ($iloLocalAccountsParam)
				{
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0} `' -f $iloLocalAccountsParam) -isVar $False	-indentlevel 19))
				}
				if ($iloDirectoryGroupsParam)
				{
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0} `' -f $iloDirectoryGroupsParam) -isVar $False	-indentlevel 19))
				}

				if ($iloHostNameParam)
				{
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0} `' -f $iloHostNameParam) 	-isVar $False	-indentlevel 19))
				}
				foreach ($_param in $_dirParamArr)
				{
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0} `' -f $_param) -isVar $False	-indentlevel 19))
				}
				foreach ($_param in $_kmParamArr)
				{
					[void]$scriptCode.Add((Generate-CustomVarCode -Prefix ('{0} `' -f $_param) -isVar $False	-indentlevel 19))
				}

				newLine    # to end the command
				$iloParam 				+= ' -IloSettings $iloPolicy '
			}


			# ############### advanced Settings /WWNN/SN/iSCSI

	}

		# Ensure there is No space afer backtick
		newLine
		if ($spTemplateParam )  # SP with a template
		{
			$prefix 	= $newCmd + ' 		 -Name $name {0}{1}{2} `' -f $spTemplateParam, $hwParam, $descParam
			[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix $prefix -isVar $False -indentlevel 1))
			newLine   # to end the command
		}
		else # SPT or standalone profile
		{
			$prefix 	= $newCmd + '	 	 -Name $name {0}{1}{2}{3}{4} `' -f $descParam, $spDescParam, $shtParam, $egParam, $affinityParam
			[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix $prefix -isVar $False -indentlevel 1))

			if ($fwParam)
			{
				$prefix		= '{0} `' -f $fwParam
				[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix $prefix -isVar $False -indentlevel 9))
			}
			if ($bmParam)
			{
				$prefix		= '{0} `' 	-f $bmParam
				[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix $prefix -isVar $False -indentlevel 9))
			}
			if ($boParam)
			{
				$prefix		= '{0} `' 	-f $boParam
				[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix $prefix -isVar $False -indentlevel 9))
			}
			if ($biosParam)
			{
				$prefix		= '{0} `' 	-f $biosParam
				[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix $prefix -isVar $False -indentlevel 9))
			}

			if ($connectionsParam)
			{
				$prefix		= '{0} `' 	-f $connectionsParam
				[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix $prefix -isVar $False -indentlevel 9))
			}

			if ($localStorageParam)
			{
				$prefix		= '{0} `' 	-f $localStorageParam
				[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix $prefix -isVar $False -indentlevel 9))
			}

			if ($StorageVolumeParam)
			{
				$prefix		= '{0} `' 	-f $StorageVolumeParam
				[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix $prefix -isVar $False -indentlevel 9))
			}

			if ($iloParam)
			{
				$prefix		= '{0}' 	-f $iloParam
				[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix $prefix -isVar $False -indentlevel 9))
			}

			newLine   # to end the command



			# sasLogicalJBOD

			ifBlock -condition 'if ($lsJBOD)' -indentlevel 1
				[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix '# ------ Configure saslogicalJBOD for profile' -isVar $False 	-indentlevel 2 ))
				[void]$scriptCode.Add( (Generate-CustomVarCode -Prefix 'prf' 		-Value ($getCmd + ' | where name -eq $name')     -indentlevel 2))
				ifBlock -condition 'if ($prf)' -indentlevel 2
					[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix '$lsJBOD | % { $prf.localStorage.sasLogicalJBODs += $_ }' -isVar $False 	-indentlevel 3 ))		
				endIfBlock -indentlevel 2
				[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('{0} -InputObject $prf' -f $saveCmd) -isVar $False 	-indentlevel 2 ))
			endIfBlock -indentlevel 1

		}
		# --- Scopes
		if ($scopes)
		{
			newLine
			[void]$scriptCode.Add( (Generate-CustomVarCode -Prefix 'object' -Value ($getCmd + ' | where name -eq $name') -indentlevel 1))
			generate-scopeCode -scopes $scopes -indentlevel 1

		}


		endIfBlock
		# Skip creating because resource already exists
		elseBlock
		[void]$scriptCode.Add(( Generate-CustomVarCode -Prefix ('write-host -foreground YELLOW "{0} already exists." ' -f $spName ) -isVar $False -indentlevel 1 ))
		endElseBlock

		newLine

		# ---------------- Split in different scripts for server profiles ONLY!
		if ($isSP)
		{
			if ($i -ge $MAXPROFILE)
			{
				newLine
	
				# Add disconnect and close the file
				[void]$ScriptCode.Add('Disconnect-OVMgmt')
			
				# ---------- Generate script to file
				writeToFile -code $ScriptCode -file $spfilename
				[void]$ListPS1Files.Add($spfilename)
	
				# Build a new set
				$scriptCode             	= [System.Collections.ArrayList]::new()
				$count++
				$spfilename 				= "$subdir\profileGroup$count" + ".PS1"
				connect-Composer  -sheetName $cSheet    -workBook $WorkBook -scriptCode $scriptCode

				$i = 1
			}
			
		}

	}

	if ($spList)
	{
		[void]$ScriptCode.Add('Disconnect-OVMgmt')
	}




	if ($isSp)
	{
		$ps1Files 	= $spFilename
		[void]$ListPs1Files.Add($spfilename)
	}


	 # ---------- Generate script to file
	 writeToFile -code $ScriptCode -file $ps1Files

	 return $ListPS1Files

}



# -------------------------------------------------------------------------------------------------------------
#
#       Main Entry
#
# -------------------------------------------------------------------------------------------------------------



# ---------------- Define Excel files
#
$startRow 				= 15

$allScriptCode 			= [System.Collections.ArrayList]::new()

if ($workbook)
{ 
	if  (test-path -path $workBook)
	{
		if ($scriptFolder)
		{
			$currentDir = $scriptFolder
		}
		else
		{
			$currentDir = Split-Path -Parent $MyInvocation.MyCommand.Path
		}

		$subdirList = @('Settings','Servers','Hypervisors','Networking','Storage','Facilities','Appliance')

		foreach ($dir in $subdirList)
		{
			$dir    = "$currentDir\$dir"
			if (-not (test-path -path $dir) )
			{
                write-host -ForegroundColor Cyan "--------- Creating folder $dir"
                write-host -ForegroundColor Cyan $CR
				md $dir
			}
		}

		$allScriptFile			= "$currentDir\allScripts.ps1"



		$sheetNames 	= (Get-ExcelSheetInfo -Path $workBook).Name
		$composer 		= 'OVdestination'
		$sequence 		= 1


		#----------------------------------------------
		#              OV Resources
		#----------------------------------------------
		[void]$allScriptCode.Add($CR)
		[void]$allScriptCode.Add('#----------------------------------------------')
		[void]$allScriptCode.Add('#              OV Resources 					 ')
		[void]$allScriptCode.Add('#----------------------------------------------')
		[void]$allScriptCode.Add($CR)

		# ================ Appliance folder
		$subdir         = "$currentdir\Appliance"

		# ---- Import Fw Baseline
		$sheet 			= 'firmwareBundle'
		$resource 		= 'firmware Baseline'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-firmwareBaseline -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}

		# ================ Settings folder
		$subdir         = "$currentdir\Settings"

		# ---- Import address pool
		$sheet 			= 'addressPool'
		$resource 		= 'address pool'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-addressPool -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}

		# ================ Networking folder
		$subdir         = "$currentdir\Networking"

		# ---- Import Network
		$sheet 			= 'ethernetNetwork'
		$resource 		= 'Ethernet network'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-ethernetNetwork -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}

		# ---- Import FC/COE Network
		$sheet 			= 'fcNetwork'
		$resource 		= 'FC network'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-fcNetwork -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}
		
		# ---- Import Network Set
		$sheet 			= 'networkSet'
		$resource		= 'network set'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource "
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-NetworkSet -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}


		# ---- Import LIG
		$sheet 			= 'logicalInterconnectGroup'
		$resource 		= 'logical interconnect group'
		######
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				$sheetArray 	= 'logicalInterconnectGroup|snmpConfiguration|snmpV3User|snmpTrap'.Split($SepChar)
				$sheetList 		= @()
				foreach ($s in $sheetArray)
				{
					if ($sheetNames -contains $s)
					{
						$sheetList += $s
					}
				}

				$sheetName 	= $sheetList -join $SepChar

				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheetName"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-LogicalInterconnectGroup -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}
		
		# ---- Import uplinkSet
		$sheet 			= 'uplinkSet'
		$resource 		= 'uplink set'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-UplinkSet -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}

		# ================ Storage folder
		$subdir         = "$currentdir\Storage"

		# ---- Import SAN Manager
		$sheet 			= 'sanManager'
		$resource 		= 'SAN Manager'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-SANmanager -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}


		# ---- Import Storage System
		$sheet 			= 'storageSystem'
		$resource 		= 'storage systems'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-storageSystem -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}

		# ---- Import Storage VolumeTemplate
		$sheet 			= 'storageVolumeTemplate'
		$resource 		= 'storage volume templates'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-StorageVolumeTemplate -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}

		# ---- Import Storage Volume
		$sheet 			= 'storageVolume'
		$resource 		= 'storage volumes'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-StorageVolume -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}


		# ---- Import Logical JBOD
		$sheet 			= 'logicalJBOD'
		$resource 		= 'logical JBODs'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-logicalJBOD -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}

		# ================ Servers folder
		$subdir         = "$currentdir\Servers"

		# ---- Import enclosure Group
		$sheet 			= 'enclosureGroup'
		$resource 		= 'enclosure group'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-EnclosureGroup -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}

		# ---- Import logical Enclosure 
		$sheet 			= 'logicalEnclosure'
		$resource 		= 'logical enclosure'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				# Will spilit 1 PS1 per LE
				$ListPS1Files = Import-LogicalEnclosure -sheetName $sheetName -workBook $workbook -subdir $subdir

				foreach ($ps1Files in $ListPS1Files)
				{
					write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
					add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				}
				$sequence++
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}

		# ---- Import server profile Templates 
		$sheet 			= 'profileTemplate'
		$resource 		= 'profile template'

		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				$sheetArray 	= 'profileTemplate|profileTemplateConnection|profileTemplateLOCALStorage|profileTemplateSANStorage|profileTemplateILO'.Split($SepChar)
				$sheetList 		= @()
				foreach ($s in $sheetArray)
				{
					if ($sheetNames -contains $s)
					{
						$sheetList += $s
					}
				}

				$sheetName 	= $sheetList -join $SepChar 
				

				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"

				$sheetName  	= "$composer|$sheetName"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				$ListPS1Files =  Import-ProfileorTemplate -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files -isSpt $True -subdir $subdir
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}

		# ---- Import server profiles 
		$sheet 			= 'profile'
		$resource 		= 'profile'

		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				$sheetArray 	= 'profile|profileConnection|profileLOCALStorage|profileSANStorage|profileILO'.Split($SepChar)
				$sheetList 		= @()
				foreach ($s in $sheetArray)
				{
					if ($sheetNames -contains $s)
					{
						$sheetList += $s
					}
				}

				$sheetName 	= $sheetList -join $SepChar 

				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"

				$sheetName  	= "$composer|$sheetName"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				$ListPS1Files 	= Import-ProfileorTemplate -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files -isSpt $False -subdir $subdir
				foreach ($ps1Files in $ListPS1Files)
				{
					write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
					add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				}
				$sequence++
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}			
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}

		# ================ Networking folder
		$subdir         = "$currentdir\Networking"

		# ---- Import logical switch group 
		$sheet 			= 'logicalSwitchGroup'
		$resource 		= 'logical switch group'

		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"

				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
						
				Import-logicalSwitchGroup -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Create $resource script"
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}

		# ---- Import logical switch
		$sheet 			= 'logicalSwitch'
		$resource 		= 'logical switch'

		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"

				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
						
				Import-logicalSwitch -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Create $resource script"
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}
		


		

		[void]$allScriptCode.Add($CR)
		[void]$allScriptCode.Add('#----------------------------------------------')
		[void]$allScriptCode.Add('#              OV Settings					 ')
		[void]$allScriptCode.Add('#----------------------------------------------')
		[void]$allScriptCode.Add($CR)

		# ================ Appliance folder
		$subdir         = "$currentdir\Appliance"
		# ---- Import users
		$sheet 			= 'user'
		$resource 		= 'OneView users and Remote Support contacts'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-user -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Create $resource script"
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}		
		}
		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}

		# ================ Facilities folder
		$subdir         = "$currentdir\Facilities"

		# ---- Import Data center
		$sheet 			= 'dataCenter'
		$resource 		= 'data center'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-dataCenter -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Create $resource script"
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}		
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}

		# ---- Import Racks 
		$sheet 			= 'racks'
		$resource 		= 'racks'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-rack -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Create $resource script"
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}		
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}



		# ================ Settings folder
		$subdir         = "$currentdir\Settings"	

		# ---- Import Proxy
		$sheet 			= 'proxy'
		$resource 		= 'proxy'

		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-proxy -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Create $resource script"
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}

		# ---- Import Time and Locale
		$sheet 			= 'timeLocale'
		$resource 		= 'time & locale'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-TimeLocale -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Create $resource script"
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}			
		
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}

		# ---- Import backup config
		$sheet 			= 'backupConfig'
		$resource 		= 'backup config'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-backup -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Create $resource script"
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}		
		}

		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}

		# ---- Import repository
		$sheet 			= 'repository'
		$resource 		= 'external repository'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-repository -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Create $resource script"
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}		
		}
		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}

		# ---- Import scope
		$sheet 			= 'scope'
		$resource 		= 'scope'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-scope -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Create $resource script"
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}
		}
		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}


		# ---- Import snmp users
		$sheet 			= 'snmpV3User'
		$resource 		= 'snmp Users'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-snmpV3User -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Create $resource script"
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}
		}
		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}

		# ---- Import snmp trap
		$sheet 			= 'snmpTrap'
		$resource 		= 'snmp traps'
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				write-host -ForegroundColor Cyan $CR
				write-host -ForegroundColor Cyan "--------- Script to import $resource"
				$sheetName  	= "$composer|$sheet"
				$ps1Files       = "$subdir\$sheet.ps1"
				
				Import-snmpTrap -sheetName $sheetName -workBook $workbook -ps1Files $ps1Files 
				write-host -ForegroundColor Cyan "Script is created ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Create $resource script"
			}
			else 
			{
				write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
			}
		}
		else 
		{
			write-host -ForegroundColor Yellow "`n No data in Excel sheet '$sheet' to create script "
		}

		[void]$allScriptCode.Add($CR)
		[void]$allScriptCode.Add('#----------------------------------------------')
        [void]$allScriptCode.Add('#              TBD -OV Appliance configuration		 ')
		[void]$allScriptCode.Add('#----------------------------------------------')
		[void]$allScriptCode.Add($CR)




		# ----- Generate all Scripts file
		write-host -ForegroundColor Cyan "`n`n--------- All-in-one script"
		write-host -ForegroundColor CYan "`n$allScriptFile contains all individual scripts that can be run to configure your new environment.`n`n"
		writeToFile -code $allScriptCode -file $allScriptFile
	}
	else 
	{
    	write-host -ForegroundColor Yellow "Excel workbook $workbook not found. Skip importing......"     
	}
}
else 
{
    write-host -ForegroundColor Yellow ' No Excel workbook provided. Skip importing......'    
}






