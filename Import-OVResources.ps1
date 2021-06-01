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
$YMLMAXLEN 		= 50
$indentDataStart = 3

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
[Hashtable]$ICModuleTypes               = @{
	"VirtualConnectSE40GbF8ModuleforSynergy"   =  "SEVC40f8" ;
	"VirtualConnectSE100GbF32ModuleforSynergy" =  "SEVC100f32" ;
	"Synergy50GbInterconnectLinkModule"        =  "SE50ILM";
	"Synergy20GbInterconnectLinkModule"        =  "SE20ILM";
	"Synergy10GbInterconnectLinkModule"        =  "SE10ILM";
	"VirtualConnectSE16GbFCModuleforSynergy"   =  "SEVC16GbFC";
	"VirtualConnectSE32GbFCModuleforSynergy"   =  "SEVC32GbFC";
	"Synergy12GbSASConnectionModule"           =  "SE12SAS";
	"571956-B21"                               =  "FlexFabric";
	"455880-B21"                               =  "Flex10";
	"638526-B21"                               =  "Flex1010D";
	"691367-B21"                               =  "Flex2040f8";
	"572018-B21"                               =  "VCFC20";
	"466482-B21"                               =  "VCFC24";
	"641146-B21"                               =  "FEX"
}

[Hashtable]$ICModuleNameTypes               = @{
	"SEVC40f8" 									= "Virtual Connect SE 40Gb F8 Module for Synergy";
	"SEVC100f32" 								= "Virtual Connect SE 100Gb F32 Module for Synergy" ;
	"SE50ILM"									= "Synergy 50Gb Interconnect Link Module";
	"SE20ILM"									= "Synergy 20Gb Interconnect Link Module";
	"SE10ILM"									= "Synergy 10Gb Interconnect Link Module";
	"SEVC16GbFC"								= "Virtual Connect SE 16Gb FC Module for Synergy";
	"SEVC32GbFC"								= "Virtual Connect SE 32Gb FC Module for Synergy"
}

[Hashtable]$FabricModuleTypes           = @{
	"VirtualConnectSE40GbF8ModuleforSynergy"    =  "SEVC40f8" ;
	"VirtualConnectSE100GbF32ModuleforSynergy"  =  "SEVC100f32" ;
	"Synergy12GbSASConnectionModule"            =  "SAS";
	"VirtualConnectSE16GbFCModuleforSynergy"    =  "SEVCFC";
	"VirtualConnectSE32GbFCModuleforSynergy"    =  "SEVCFC";
}

[Hashtable]$ICModuleToFabricModuleTypes = @{
	"SEVC40f8"                                  = "SEVC40f8" ;
	"SEVC100f32"                                = "SEV100f32" ;
	'SE50ILM'                                   = "SEV100f32" ;
	'SE20ILM'                                   = "SEVC40f8" ;
	'SE10ILM'                                   = "SEVC40f8" ;
	"SEVC16GbFC"                                = "SEVCFC" ;
	"SEVC32GbFC"                                = "SEVCFC" ;
	"SE12SAS"                                   = "SAS"
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

[Hashtable]$YMLconsistencyCheckingSptEnum  			= @{
	'None'			= 'UnChecked';
	'Exact'			= 'Checked';
	'Minimum'		= 'CheckedMinimum'
	}
[Hashtable]$YMLconsistencyCheckingLigEnum  			= @{
	'None'			= 'NotChecked';
	'Exact'			= 'ExactMatch';
	'Minimum'		= 'CheckedMinimum'
	}

[HashTable]$iLOPrivilegeParamEnum     = @{
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

[HashTable]$YMLiLOPrivilegeParamEnum     = @{
	'userConfigPriv'			= 'userConfigPriv           : True,'  	;
	'remoteConsolePriv'			= 'remoteConsolePriv        : True,'  	;
	'virtualMediaPriv'			= 'virtualMediaPriv         : True,'  	;
	'virtualPowerAndResetPriv'	= 'virtualPowerAndResetPriv : True,'	;
	'iloConfigPriv'				= 'iLOConfigPriv            : True,'	;
	'loginPriv'					= 'loginPriv                : True,'	;
	'hostBIOSConfigPriv'		= 'hostBIOSConfigPriv       : True,'	;
	'hostNICConfigPriv'			= 'hostNICConfigPriv        : True,'	;
	'hostStorageConfigPriv'		= 'hostStorageConfigPriv    : True,'			 
}

[HashTable]$YMLiLOdirAuthParamEnum     = @{
	'DirectoryDefault'			= 'defaultSchema';  
	'HPEExtended'				= 'extendedSchema'; 
	'Disabled'					= 'disabledSchema'
}

[HashTable]$YMLsnmpSecurityLevelEnum     = @{
	'AuthOnly'					= 'Authentication';  
	'AuthAndPriv'				= 'Authentication and privacy';
	'None'						= 'None'
}

[HashTable]$YMLsnmpProtocolEnum     = @{
	'des56'						= 'DES';  
	'3des'						= '3DES';
	'AuthOnly'					= 'AES';  
	'aes128'					= 'AES-128';
	'aes256'					= 'AES-256';  
	'None'						= 'None'
}


[HashTable]$YMLtype540Enum 			= @{
	'ethernet'					= 'ethernet-networkV4';
	'fcnetwork'					= 'fc-networkV4';
	'networkset'				= 'network-setV5';
	'ethernetSettings'			= 'EthernetInterconnectSettingsV6'   ;
	'serverProfileTemplate'		= 'ServerProfileTemplateV7'			 ;
	'serverProfile'				= 'ServerProfileV11'			     
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
	[string]$fabricModuleName
	[string]$logicalPortConfigInfos
	[string]$fcUplinkSpeed 
	[string]$loadBalancingMode
	[string]$lacpTimer
	[string]$primaryPort
	[string]$privateVLanDomains
	[string]$consistencyChecking
} 

class uplConfig {
	[int]$frame
	[int]$bay
	[string]$bayModule
	[string]$bayModuleName
	[string]$portName
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
	[string]$enclosureNewName
	[string]$enclosureGroup
	[string]$manualAddresses
	[string]$firmwareBaseline	
	[Boolean]$forceInstallFirmware	
	[string]$logicalInterconnectUpdateMode	
	[Boolean]$updateFirmwareOnUnmanagedInterconnect	
	[Boolean]$validateIfLIFirmwareUpdateIsNonDisruptive
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
	[string]$bootTarget
	[string]$targetLUN 					# LunID

	[boolean]$userDefined = $False
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

function build-portInfos ($bayconfig,$logicalPortInfo )
{
	$bayPortInfos 	= [System.Collections.ArrayList]::new()
	if ( ($bayConfig) -and ($bayConfig -notlike '*SAS*') )
	{	
		$bayConfig 		= $bayConfig.replace($Space, '')

		$bayConfigArr 	= $bayConfig.Split($CR)
		foreach ($_bConfig in $bayConfigArr)
		{
			$_frame, $_bay  	= $_bConfig.Split('{')											# Each frame looks like Frame1={Bay2='SEVC40f8'|Bay5='SE20ILM'}
			if ($_bay)
			{
				$_bay				= $_bay.Trim() -replace '{','' -replace '}', ''
				$_bayArr 			= $_bay.Split($SepChar)
				foreach($_b in $_bayArr)
				{
					$_bayElement 		= new-object -typeName uplConfig
					$_bayElement.frame 	= [int]($_frame.replace('Frame','').replace('=','').Trim())
					$_bayNumber, $_ICmodule			= $_b.Split('=')
					$_bayElement.bay				= [int]($_bayNumber.replace('Bay', ''))
					$_ICmodule						= $_ICmodule.trim() -replace "'", ''
					$_bayElement.bayModule 			= $_ICmodule
					$_bayElement.bayModuleName  	= $ICModuleNameTypes.item($_ICmodule)

					[void]$bayPortInfos.Add($_bayElement)

				}
			}
		}

		# ------- logical port info 'Enclosure1:Bay2:Q1|Enclosure2:Bay5:Q1.1'
		if ($logicalPortInfo)
		{
			$portInfoArr = $logicalPortInfo.Split($SepChar)
			foreach ($_p in $portInfoArr)
			{
				$_encl,$_bay,$_port = $_p.Trim().Split(':')
				$_frameNumber 	= $_encl[-1]
				$_bayNumber		= $_bay[-1]
				$_portNumber 	= $_port.replace('.', ':') 	# normalize port naming convention	 
				for ( $index = 0; $index -lt $bayPortInfos.count; $index++ )
				{
					$_el 	= $bayPortInfos[$index]
					if (($_el.frame -match $_frameNumber) -and ($_el.bay -match $_bayNumber))
					{
						$_el.portName 	= $_portNumber
					}
				}


			} 
		}
	}
	return $bayPortInfos
}
Function Generate-PSCustomVarCode ([String]$Prefix, [String]$Suffix, [String]$Value, $indentlevel = 0, $isVar = $True)
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

function YMLwriteToFile ([System.Collections.ArrayList]$code,[string]$file)
{
	if ($code.Count -gt 0)
	{
		$code | out-file $file
	}
}

# ---------- Internal Helper funtion
function generate-scopeCode( $scopes, $indentlevel = 0)
{
	$arr 		= [system.Collections.ArrayList]::new()
	$arr 		= $scopes.split($SepChar) | % { '"{0}"' -f $_ }
	$scopeList  = '@({0})' -f [string]::Join($Comma, $arr)

	$scopeCode	= ($TAB * $indentLevel) + ('{0}' -f $scopeList)  + ' | % { get-OVScope -name $_ | Add-OVResourceToScope -InputObject $object }'

	[void]$PSscriptCode.Add($scopeCode)
}

function startBlock($indentlevel = 0, $code = $PSscriptCode )
{
	[void]$code.Add(($TAB * $indentLevel) + '{')
}


function endBlock($indentlevel = 0)
{
	[void]$PSscriptCode.Add(($TAB * $indentLevel) + '}')
}

function ifBlock($condition, $indentlevel = 0, $code = $PSscriptCode )
{

	[void]$code.Add((Generate-PSCustomVarCode -Prefix  $condition 					-isVar $False -indentlevel $indentlevel) )
	[void]$code.Add(($TAB * $indentLevel) + '{')
}

function endIfBlock ($condition='', $indentlevel = 0, $code = $PSscriptCode)
{

	[void]$code.Add(($TAB * $indentLevel) + '}' + ' # {0}' -f $condition)
}

function elseBlock($indentlevel = 0, $code = $PSscriptCode)
{
	[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'else' 						-isVar $False -indentlevel $indentlevel) )
	[void]$code.Add(($TAB * $indentLevel) + '{')
}


function endElseBlock ($indentlevel = 0, $code = $PSscriptCode)
{
	[void]$code.Add(($TAB * $indentLevel) + '}')
}

function newLine($code = $PSscriptCode)
{
	[void]$code.Add($CR)
}

function add_to_allScripts ($text , $ps1Files)
{
	[void]$allPSscriptCode.Add($text)
	[void]$allPSscriptCode.add($ps1Files)
	[void]$allPSscriptCode.add($CR)
}

# ----------
# ---------- YML functions
Function Generate-YMLCustomVarCode ([String]$Prefix, [String]$Suffix, [String]$Value, $indentlevel = 0, $isVar = $False, $isItem = $False, $distance = $YMLMAXLEN)
{

    $TAB        = $Space  * 3 



    if ($isVar)
    {
        $Prefix = '- {0}' -f $Prefix
    }
	else 
	{
		$Prefix = '  {0}' -f $Prefix	
	}



    $VarName    = ($TAB * $indentlevel) + ('{0}{1}' -f $Prefix, $Suffix)
    

    $len 	= $VarName.Length
    $len 	= if ($len -ge $distance) {$distance} else { $distance - $len}
    $pad    = $Space * $len

    if ($isItem)
    {
        $out 	= '{0}{1}{2}' -f $VarName,$pad, $Value
    }
    else
    {
        $out 	= '{0}:{1}{2}' -f $VarName,$pad, $Value
    }

    if ($prefix -like '*--*')
    {
        $out    = $prefix.trim()
    }
	

	
    return $out
}


function Generate-ymlheader ([string]$title, $code = $YMLscriptCode )
{
    [void]$code.Add((Generate-YMLCustomVarCode -prefix '---' ))                                    
    [void]$code.Add((Generate-YMLCustomVarCode -prefix name      -value $title                   -isVar $True))
    [void]$code.Add((Generate-YMLCustomVarCode -prefix hosts     -value 'localhost' ))             
    [void]$code.Add((Generate-YMLCustomVarCode -prefix vars))                                  
	[void]$code.Add((Generate-YMLCustomVarCode -prefix config    -value "'oneview_config.json'"  -indentlevel 1))
	[void]$code.Add((Generate-YMLCustomVarCode -prefix variant   -value Synergy  				 -indentlevel 1))
    [void]$code.Add((Generate-YMLCustomVarCode -prefix tasks  ))
}

function Generate-ymlTask ([string]$Title, [string]$comment,[string]$ovTask, [string]$state = 'present', $isData = $True, $iseTag = $False,  $code = $YMLscriptCode )
{
	newLine	-code $YMLscriptCode
		 
	[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix $comment  						-isItem $True  					-indentlevel 1 ))
	   
	[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 			-prefix name  				-value $title	-isVar $True 	-indentlevel 1 ))
	[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 			-prefix $ovTask		 										-indentlevel 1 ))
	[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 			-prefix config				-value "'{{config}}'"			-indentlevel 2 ))
	if ($iseTag)
	{
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 		-prefix validate_etag		-value False					-indentlevel 2 ))
	}
	if ($isData)
	{
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 		-prefix state				-value $state					-indentlevel 2 ))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 		-prefix 'data'												-indentlevel 2 ))
	}

	   
}

# ----- connect Composer
function connect-Composer([string]$sheetName, [string]$WorkBook, $PSscriptCode )
{
	$composer 				= get-datafromSheet -sheetName $sheetName -workbook $WorkBook


	$hostName 			= $composer.name
	$userName 			= $composer.Username
	$password 			= $composer.password
	$authDomain 		= if ($NULL -ne $composer.authenticationDomain) {$composer.authenticationDomain} else {'LOCAL'}

	

	newLine
	[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'write-host "`n"'  -isVar $False )) 
	[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Connecting to OneView {0}" ' -f $hostName) -isVar $False ))
	
	ifBlock 	-condition 'if ($global:ConnectedSessions)'
	[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '$d = disconnect-OVMgmt -ApplianceConnection $global:ConnectedSessions' -isVar $false -indentLevel 1))
	endIfBlock 
	generate-credentialCode -username $userName -password $password -component 'OneView' -PSscriptCode $PSscriptCode
	[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('connect-OVMgmt -hostname {0} -credential $cred -loginAcknowledge:$True -AuthLoginDomain "{1}" ' -f $hostName,$authDomain) -isVar $False ))
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
Function generate-credentialCode ($username, $password, $component,$indentLevel=0, $PSscriptCode)
{
	[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'userName' -value ( ' "{0}" ' -f $userName ) -indentlevel $indentLevel))
	[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'password' -value ( ' "{0}" ' -f $password ) -indentlevel $indentLevel))

	ifBlock -condition 'if ($password)' -indentlevel $indentLevel
	[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'securePassword' -value '$password | ConvertTo-SecureString -AsPlainText -Force' -indentLevel ($indentLevel + 1) ))
	endIfBlock 	-indentlevel $indentLevel 

	elseBlock 	-indentlevel $indentLevel 
	$value 		= 'Read-Host "{0}: enter password for user {1}" -AsSecureString ' -f $component, $username
	[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'securePassword' -Value $value  -indentlevel ($indentLevel+1)))
	endElseBlock  -indentlevel $indentLevel

	[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'cred ' -Value 'New-Object System.Management.Automation.PSCredential  -ArgumentList $userName, $securePassword'  -indentlevel $indentLevel))
	newLine
	
}

# ---------- firmware Bundle
Function Import-firmwareBaseline([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $PSscriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  

    foreach ($fw in $List)
    {
		$filename 			= $fw.filename
		$name 				= $fw.name
		if ($filename)
		{
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Adding firmware Baseline {0} "' -f $name) -isVar $False ))
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('Add-OVBaseline -file  "{0}"' -f $filename) -isVar $False ))
            newLine
        }
	}

	if ($List)
	{
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	}

	# ---------- Generate script to file
	writeToFile -code $PSscriptCode -file $ps1Files
}


# ---------- Data Center
Function Import-dataCenter([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $PSscriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	
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
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating datacenter {0} "' -f $name) -isVar $False ))
			$value 					= "Get-OVDataCenter -ErrorAction SilentlyContinue -Name  '{0}' " -f $name #HKD
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'dc' 						-Value $value ))		

			ifBlock -condition 'if ($dc -eq $Null)'
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('# -------------- Attributes for dc {0} ' -f $name) -isVar $False -indentlevel 1))
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'name' 			-Value ("'{0}'" -f $name) -indentlevel 1))

			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'width' 			-Value ('{0}' 	-f $width) -indentlevel 1))
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'depth' 			-Value ('{0}' 	-f $depth) -indentlevel 1))
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'millimeters' 	-Value ('${0}' 	-f $millimeters) -indentlevel 1))

			$nameParam 			= ' -Name "{0}" '		-f $name
			$widthParam 		= ' -Width {0}  ' 		-f $width
			$depthParam 		= ' -Depth {0}  ' 		-f $depth
			$millimetersParam 	= ' -Millimeters:${0}'	-f $millimeters

			$voltageParam  		= $null
			if ($defaultVoltage)
			{
				$voltageParam 	= ' -DefaultVoltage $defaultVoltage'
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'defaultVoltage' 	-Value ('{0}' -f $defaultVoltage) -indentlevel 1))

			}

			$currencyParam  		= $null
			if ($currency)
			{
				$currencyParam 	= ' -Currency $currency'
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'currency' 		-Value ("'{0}'" -f $currency) -indentlevel 1))

			}

			$powerCostsParam  		= $null
			if ($powerCosts)
			{
				$powerCostsParam 	= ' -PowerCosts $powerCosts'
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'powerCosts' 		-Value ('{0}' -f $powerCosts) -indentlevel 1))

			}

			$coolingCapacityParam  = $null
			if ($coolingCapacity)
			{
				$coolingCapacityParam 	= ' -CoolingCapacity $coolingCapacity'
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'coolingCapacity' -Value ('{0}' -f $coolingCapacity) -indentlevel 1))
			}

			$address1Param  = $null
			if ($address1)
			{
				$address1Param 	= ' -Address1 $address1'
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'address1' 		-Value ("'{0}'" -f $address1) -indentlevel 1))
			}

			$address2Param  = $null
			if ($address2)
			{
				$address2Param 	= ' -Address2 $address2'
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'address2' 		-Value ("'{0}'" -f $address2) -indentlevel 1))
			}

			$cityParam  = $null
			if ($city)
			{
				$cityParam 	= ' -City $city'
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'city' 			-Value ("'{0}'" -f $city) -indentlevel 1))
			}

			$stateParam  = $null
			if ($state)
			{
				$stateParam 	= ' -State $state'
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'state' 			-Value ("'{0}'" -f $state) -indentlevel 1))
			}

			$postCodeParam  = $null
			if ($postCode)
			{
				$postCodeParam 	= ' -PostCode $postCode'
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'postCode' 		-Value ("'{0}'" -f $postCode) -indentlevel 1))
			}

			$countryParam  = $null
			if ($country)
			{
				$postCodeParam 	= ' -Country $country'
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'country' 		-Value ("'{0}'" -f $country) -indentlevel 1))
			}

			$timezoneParam  = $null
			if ($timezone)
			{
				$timezoneParam 	= ' -TimeZone $timezone'
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'timezone' 		-Value ("'{0}'" -f $timezone) -indentlevel 1))
			}

			$primaryContactParam  = $null
			if ($primaryContact)
			{
				$primaryContactParam 	= ' -PrimaryContact $primaryContact'
				$value					= 'Get-OVRemoteSupportContact -Name "{0}" ' -f $primaryContact
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'primaryContact' 	-Value  $value -indentlevel 1))
			}

			$secondaryContactParam  = $null
			if ($secondaryContact)
			{
				$secondaryContactParam 	= ' -SecondaryContact $secondaryContact'
				$value					= 'Get-OVRemoteSupportContact -Name "{0}" ' -f $secondaryContact
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'secondaryContact' 	-Value $value  -indentlevel 1))
			}
			

			# Code
			# ensure that there is no space after backstick(`)
			newLine

			$prefix		= 'New-OVDatacenter {0}{1}{2}{3} `' -f $nameParam,$widthParam,$depthParam,$millimetersParam
			[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix $prefix -isVar $False -indentLevel 1)) 

			$prefix 	= '{0}{1}{2}{3} `' 	-f $voltageParam, $powerCostsParam, $currencyParam, $coolingCapacityParam 
			[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix $prefix -isVar $False -indentLevel 2))

			$prefix 	= '{0}{1}{2}{3}{4}{5}{6} `' 	-f $address1Param, $address2Param, $cityParam, $stateParam, $postCodeParam,  $countryParam , $timezoneParam
			[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix $prefix -isVar $False -indentLevel 2))

			$prefix 	= '{0}{1} `' 	-f $primaryContactParam, $secondaryContactParam
			[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix $prefix -isVar $False -indentLevel 2))

			newLine # to end the command
			endifBlock -condition 'if ($dc -eq $Null)'

			# Skip creating because resource already exists
			elseBlock
			[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $name + ' already exists.') -isVar $False -indentlevel 1 ))
			endElseBlock


		newLine
		}

	}

	if ($List)
	{
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	}

	# ---------- Generate script to file
	writeToFile -code $PSscriptCode -file $ps1Files
}


# ---------- Rack
Function Import-rack([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $PSscriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  


	foreach ($rack in $List)
	{
		$dcName 			= $rack.name
		$rackSN 			= $rack.rackSerialNumber
		$x 					= $rack.xCoordinate
		$y 					= $rack.yCoordinate
		$millimeters		= $rack.millimeters

		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Adding rack {0} to datacenter {1} "' -f $rackSN, $dcName) -isVar $False ))
		newLine

		$value 				= "Get-OVDataCenter -ErrorAction SilentlyContinue -Name $dcName "  #HKD 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'dc' 						-Value $value ))	
		$value 				= 'Get-OVRack | where serialNumber -match "{0}" ' -f $rackSN
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'rack' 					-Value $value ))
		
		newLine

		ifBlock 	-condition 'if ( ($dc -ne $Null) -and ($rack -ne $Null) )' 
		$value 				= '$rack.Uri'
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'rackUri' -Value $value -indentlevel 1 ))

		#$value 				=  '$dc.contents | where resourceUri -match $rackUri'
		#[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'rack_in_dc' -Value $value -indentlevel 1 ))

		ifBlock 	-condition 'if ($null -eq ($dc.contents | where resourceUri -match $rackUri) )'  -indentlevel 1
		# Code here
		$dcParam  			= ' -DataCenter $dc ' 
		$inputParam 		= ' -InputObject $rack '
		$coordParam 		= ' -X {0} -Y {1} -Millimeters:${2}'		-f $x, $y, $millimeters
		$prefix 			= 'Add-OVRackToDataCenter {0}{1}{2} ' 	-f $inputParam,$dcParam, $coordParam
		
		[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix $prefix -isVar $False -indentLevel 2)) 
		endIfBlock 	-indentlevel 1

		# Rack already defined in dc
		elseBlock 	-indentlevel 1
		[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $rackSN + ' already defined in data center.') -isVar $False -indentlevel 2 ))
		endElseBlock -indentlevel 1

		endIfBlock 	-condition '$dc -ne $Null and $rack $ne $Null' 

		# Data center not existed
		elseBlock
		[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW ' + "datacenter '{0}'" -f $dcName + " or rack '{0}'" -f $rackSN  + ' do not exist. Define datacenter first') -isVar $False -indentlevel 1 ))
		endElseBlock



		# Code
		# ensure that there is no space after backstick(`)
		newLine



	}

	if ($List)
	{
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	}

	# ---------- Generate script to file
	writeToFile -code $PSscriptCode -file $ps1Files
}


# ---------- proxy
Function Import-proxy([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $PSscriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  


	foreach ($proxy in $List)
	{
		$name 				= $proxy.server
		$protocol 			= $proxy.protocol
		$port 				= $proxy.port
		$username 			= $proxy.Username

		
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Importing proxy "' -f $name) -isVar $False ))

		$hostNameParam 		= ' -Hostname "{0}"' 	-f $name
		$https 				= if ($protocol -eq 'https') {'$True'} else {'$False'}
		$httpsParam 		= ' -Https:{0}' 		-f $https
		$portParam 			= ' -Port {0}'			-f $port

		$userParam 			= $null
		if ($username)
		{
			$value 			= 'Read-Host "Proxy Setting: enter password for user {0}" -AsSecureString ' -f $username
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'securepass' -value  $value))
			$userParam 		= ' -Username "{0}" -password $securepass ' -f $username
		}
		# Code
		# ensure that there is no space after backstick(`)
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('Set-OVApplianceProxy {0}{1}{2}{3}' -f $hostNameParam,$portParam,$httpsParam,$userParam) -isVar $False ))
		newLine

	}

	if ($List)
	{
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	}
	# ---------- Generate script to file
	writeToFile -code $PSscriptCode -file $ps1Files
}


# ---------- proxy
Function Import-TimeLocale([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $PSscriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	
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
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('Set-OVApplianceDateTime -locale "{0}" -ntpServers {1} {2} {3}' -f $locale,$ntpServers,$syncParam,$pollParam) -isVar $False ))
		newLine

	}

	if ($List)
	{
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	}

	# ---------- Generate script to file
	writeToFile -code $PSscriptCode -file $ps1Files
}


 Function Import-backup([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
 {
	 $PSscriptCode            = [System.Collections.ArrayList]::new()
	 $cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	 connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	 
	 $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	 if ($List)
	 {
		 [void]$PSscriptCode.Add($codeComposer)
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
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'hostSSHKey' -value ("'{0}'" -f $publicKey) ))

				if ($NULL -eq $password)
				{
					$value 			= 'Read-Host "Backup Config Setting --> enter password for user {0}" -AsSecureString ' -f $username
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'securepass' -value  $value))
				}
				else
				{
					$value 			= "'$password' | ConvertTo-SecureString -AsPlainText -Force "
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'securepass' -value  $value))
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
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('Set-OVAutomaticBackupConfig {0} `' -f $remoteParam) -isVar $False ))
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0} {1}' -f $userParam, $scheduleParam) -isVar $False -indentlevel 1))
			newLine
		}
		else # Disable backup
		{
			# Code
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'Set-OVAutomaticBackupConfig  -Disabled:$True '  -isVar $False ))
		}

	 }
	 
	 if ($List)
	 {
		 [void]$PSscriptCode.Add('Disconnect-OVMgmt')
	 }
  # ---------- Generate script to file
  writeToFile -code $PSscriptCode -file $ps1Files
 }


 Function Import-repository([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
 {
	 $PSscriptCode             = [System.Collections.ArrayList]::new()
	 $cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	 connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	 
	 $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	 if ($List)
	 {
		 [void]$PSscriptCode.Add($codeComposer)
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
		

		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating external repository {0} "' -f $name) -isVar $False ))

		if ($username)
		{
			generate-credentialCode -username $username -password $password -component 'REPOSITORY' -PSscriptCode $PSscriptCode
			$credParam 		= ' -Credential $cred'
	 	}

		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('New-OVExternalRepository -Name "{0}" {1}{2}{3}{4}{5}' -f $name, $hostParam, $httpParam, $dirParam, $credParam, $certParam) -isVar $False ))
	

		# Code
	 }
	 
	 if ($List)
	 {
		 [void]$PSscriptCode.Add('Disconnect-OVMgmt')
	 }
	# ---------- Generate script to file
	writeToFile -code $PSscriptCode -file $ps1Files
 }

 
 Function Import-scope([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
 {
	 $PSscriptCode             = [System.Collections.ArrayList]::new()
	 $cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	 connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	 
	 $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	 if ($List)
	 {
		 [void]$PSscriptCode.Add($codeComposer)
	 }
	 
   
	 foreach ($scope in $List) 
	 {
		$name 				= $scope.name
		$description		= $scope.description
		$resource 			= $scope.resource



		$descParam 			= if ($description) { ' -Description "{0}" ' -f $description} else {''}
		
		newLine
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating scopes {0} "' -f $name) -isVar $False ))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix username -Value ("'{0}'" -f $name)  ))

		ifBlock 		-condition ('if ($null -eq (get-OVScope -name "{0}" ))' -f $name) #HKD
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('New-OVScope -name "{0}" {1} ' -f $name, $descParam) -isVar $False -indentlevel 1))
		
		newLine
		
		if ($resource)
		{
			$resArray  		= $resource.Split($SepChar)
			$index 			= 1
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix resources -value '@()' -indentLevel 1) )
			foreach ($res in $ResArray)
			{
				$resType, $resName 	= $res.Split(';')
                $resType 			= $resType.replace('type=', '').Trim()
                if ($resName)
				{
                    $resName 			= '"' + $resName.replace('name=', '').Trim() + '"'  # etract name and surround with quotes
                }

				$value 	= "Get-OV$resType -name '$resName' " # HKD
				$prefix = "res$Index"
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix $prefix -value $value -indentlevel 1 ))

				ifBlock		-condition ('if (${0})' -f $prefix) -indentlevel 1
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix resources -value ('$resources + ${0}' -f $prefix) -indentlevel 2 ))
				endIfBlock -indentlevel 1 

				ifBlock 	-condition 'if ($resources)'
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('get-OVScope -name "{0}" | Add-OVResourceToScope -InputObject $resources ' -f $name) -isVar $False -indentlevel 2))
				endIfBlock -indentlevel 1
				newLine
			}


		}
		
		
		endIfBlock		-condition '$null -eq (get-OVScope....'
		# Skip creating because resource already exists

		elseBlock
		[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $name + ' already exists.') -isVar $False -indentlevel 1 ))
		endElseBlock


	 }
	 
	if ($List)
	{
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	}
   # ---------- Generate script to file
   writeToFile -code $PSscriptCode -file $ps1Files
 }


 Function Import-snmpTrap([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
 {
	 $PSscriptCode             = [System.Collections.ArrayList]::new()
	 $cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	 
	 
	 $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	 if ($List)
	 {
		$List 				= $List | where source -eq 'Appliance' # Select Appliance snmp user only
		connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
		create-snmpTrap -list $List -isSnmpAppliance $True
		newLine
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	 }
   # ---------- Generate script to file
   writeToFile -code $PSscriptCode -file $ps1Files
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
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'snmpV3user' 	-value ("Get-OVSnmpV3User -Name '$snmpV3User' " ) -indentlevel $indent))	# HKD
			}
			else 
			{
				# Use snmpV3Users / snmpV3UserNames variable in create_snmpV3user
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'this_index' -Value ('[array]::IndexOf($snmpV3Usernames, "{0}" )' -f $snmpV3User) -indentlevel $indent))	
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'snmpV3user' -Value '$snmpV3Users[$this_index]' -indentlevel $indent								  ))
			}

			$snmpV3userParam = ' -SnmpV3user $snmpV3user '
		}

		$new_trap		= '$trap{0}' -f $Index	
		$trapCmd 		= if ($isSnmpAppliance) { 'New-OVApplianceTrapDestination '} else {'New-OVSnmpTrapDestination '}

		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ("write-host -foreground CYAN '----- Creating snmp trap {0} '"  -f $new_trap)  -isVar $False -indentlevel $indent))
		[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('{0}' -f $new_trap) -value ('{0}{1}{2}{3}{4}{5}{6}{7}' -f $trapCmd, $formatParam,$destinationParam,$portParam,$trapTypeParam,$engineIdParam,$communityParam,$snmpV3userParam ) -isVar $false -indentlevel  $indent ))
		newLine
		$snmpTraps		+= $new_trap		

	 }

	 if ($isSnmpLig)
	 {
		 newLine
		 [void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix 'snmpTraps' -value ('@({0})' -f ($snmpTraps  -join $COMMA) ) -indentlevel  $indent ))
	 }	
	 

   # ---------- Generate script to file
   writeToFile -code $PSscriptCode -file $ps1Files
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
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'snmpv3Users' 	-value	'[System.Collections.ArrayList]::new()' -indentlevel $indent ) )
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'snmpv3Usernames' -value	'[System.Collections.ArrayList]::new()' -indentlevel $indent ) )
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

            [void]$PSscriptCode.Add($CR)
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating snmpV3 user {0} " ' -f $snmpV3User)  -isVar $False -indentlevel $indent))

			if ($isSnmpAppliance)
            {
				$condition              = 'if ($null -eq (Get-OVSnmpV3User -name  "{0}"))' -f $snmpV3User #HKD
            	[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix  $condition 	-isVar $False -indentlevel $indent  ) )
            	[void]$PSscriptCode.Add('{')

				$applianceSnmpParam 	= ' -ApplianceSnmpUser '
				$indent 				+= 1
			}
			$secParam 				= $null
			switch ($securityLevel)
			{
				'AuthOnly'		
					{

						[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'authPassword' -value ('"{0}" | ConvertTo-SecureString -AsPlainText -Force' -f $authPassword) -indentLevel $indent))
						$secParam	= ' -SecurityLevel "{0}" -AuthProtocol "{1}" -AuthPassword $authPassword' -f $securityLevel, $authProtocol
					}
				'AuthAndPriv'
					{
						[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'authPassword' -value ('"{0}" | ConvertTo-SecureString -AsPlainText -Force' -f $authPassword) -indentLevel  $indent ))
						[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'privPassword' -value ('"{0}" | ConvertTo-SecureString -AsPlainText -Force' -f $privPassword)  -indentLevel  $indent))
						$secParam	= ' -SecurityLevel "{0}" -AuthProtocol "{1}" -AuthPassword $authPassword -PrivProtocol "{2}" -PrivPassword $privPassword' -f $securityLevel, $authProtocol, $privProtocol 
					}										
			}

			$new_snmpv3User 		= '$snmpv3User{0}' -f $Index
			[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('{0}' -f $new_snmpv3user) -value ('new-OVSnmpV3User {0} -UserName "{1}" {2}' -f $applianceSnmpParam, $snmpV3User, $secParam ) -isVar $false -indentlevel  $indent ))
			$snmpv3Users 		+= $new_snmpv3User
			$snmpv3Usernames	+= $snmpv3User			# Collect user names

			if ($isSnmpAppliance)
            {
				[void]$PSscriptCode.Add('}')
				# Skip creating because resource already exists
            	[void]$PSscriptCode.Add('else')
            	[void]$PSscriptCode.Add('{')
				[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $snmpV3User + ' already exists.') -isVar $False -indentlevel  $indent ))
				[void]$PSscriptCode.Add('}')
			}


			
		}
	 
    
	 }
	 
	 if ($snmpv3Users -and $isSnmpLig)
	 {
		 newLine
		 [void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix 'snmpv3Users' -value ('@({0})' -f ($snmpv3Users  -join $COMMA) ) -indentlevel  $indent ))
		 
		 # Set collection of snmp v3 user names
		 $snmpv3UserNames 	= $snmpv3UserNames | % { "'{0}'" -f  $_ }  	# Add prefix $
		 [void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix 'snmpv3Usernames' -value ("@({0})" -f ($snmpv3UserNames  -join $COMMA) ) -indentlevel  $indent ))
	 }

	 

 }


 Function Import-snmpV3User([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
 {
	 $PSscriptCode          = [System.Collections.ArrayList]::new()
	 $cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	 connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	 
	 $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	 if ($List)
	 {
		 $List 				= $List | where source -eq 'Appliance' # Select Appliance snmp user only
		 [void]$PSscriptCode.Add($codeComposer)
		 create-snmpV3User -list $List -isSnmpAppliance $True
		 newLine
		 [void]$PSscriptCode.Add('Disconnect-OVMgmt')
	 }
 
   # ---------- Generate script to file
   writeToFile -code $PSscriptCode -file $ps1Files
 }


 Function Import-YMLSNMP([string]$sheetName, [string]$WorkBook, [string]$YMLfiles )
{
    $YMLscriptCode          										= [System.Collections.ArrayList]::new()
	$cSheet, $snmpConfigSheet, $snmpV3UserSheet, $snmpTrapSheet 	= $sheetName.Split($SepChar)       # composer

	[void]$YMLscriptCode.Add((Generate-ymlheader -title 'Configure SNMP for Appliance'))
	
	$snmpConfigList 		= get-datafromSheet -sheetName $snmpConfigSheet -workbook $WorkBook  	
	$snmpConfigList			= $snmpConfigList | where source -eq 'Appliance'

	$snmpV3UserList 		= get-datafromSheet -sheetName $snmpV3UserSheet -workbook $WorkBook  	
	$snmpV3UserList			= $snmpV3UserList | where source -eq 'Appliance'

	$snmpTrapList 			= get-datafromSheet -sheetName $snmpTrapSheet -workbook $WorkBook  	
	$snmpTrapList			= $snmpTrapList | where source -eq 'Appliance'


	foreach ( $_snmp in $snmpConfigList)
	{	
		# snmpV1 information
		$consistencyChecking	= $_snmp.consistencyChecking
		$communityString		= $_snmp.communityString

	
		if ($communityString)
		{
			$comment 			= '# ---------- Appliance snmp Configuration ' 				
			$title 				= ' appliance read community string'
			[void]$YMLscriptCode.Add((Generate-ymlTask 		 	-title $title -comment $comment -OVTask 'oneview_appliance_device_read_community'))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'communityString'		-Value $communityString  				-indentlevel $indentDataStart ))
		}
	

		# ----  snmpV3Users 
		foreach ($_user in $snmpV3UserList)
		{
			$userName 			= $_user.userName
			$secLevel			= $_user.securityLevel
			$authProtocol 		= $_user.authProtocol
			$authPassword		= $_user.authPassword
			$privacyProtocol 	= $_user.privacyProtocol
			$privacyPassword 	= $_user.privacyPassword

			$secLevel 			= $YMLsnmpSecurityLevelEnum.item($secLevel)

			$comment 			= '# ---------- Appliance snmp V3 User {0} ' 	-f $userName			
			$title 				= ' Create snmp v3 user {0}' 					-f $userName
			[void]$YMLscriptCode.Add((Generate-ymlTask 		 		-title $title -comment $comment -OVTask 'oneview_appliance_device_snmp_v3_users'))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-Prefix 	type 					-value Users				-indentlevel $indentDataStart))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-Prefix 	userName 				-value $userName			-indentlevel $indentDataStart))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-Prefix 	securityLevel 			-value $secLevel			-indentlevel $indentDataStart))
			if ($secLevel -like '*Privacy*')
			{
				$_privacy 		= $YMLsnmpProtocolEnum.item($privacyProtocol.Trim())
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -Prefix 	authenticationProtocol	-value $authProtocol 		-indentlevel $indentDataStart))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -Prefix 	authenticationPassphrase -value "'$authPassword'"	-indentlevel $indentDataStart))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -Prefix 	privacyProtocol 		-value $_privacy			-indentlevel $indentDataStart))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -Prefix 	privacyPassphrase		-value "'$privacyPassword'"	-indentlevel $indentDataStart))	
			}
			if ($secLevel -eq 'Authentication')
			{
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -Prefix 	authenticationProtocol	-value $authProtocol 		-indentlevel $indentDataStart))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -Prefix 	authenticationPassphrase -value "'$authPassword'"	-indentlevel $indentDataStart))
			}	
				
				

		}
	
		# ----- snmpTrap 
		foreach ($_trap in $snmpTrapList)
		{
			$format 					= $_trap.format			#snmpv1 or snmpV3
			$destinationAddress			= $_trap.destinationAddress
			$port 						= $_trap.port              
			$communityString 			= $_trap.communityString
			$snmpV3User 				= $_trap.userName

			$_format 					= $format.trim().ToLower().replace('snmp','')

			if ($snmpV3User)
			{
				$var_user 					= "var_{0}" -f ($snmpV3User.Trim().replace($Space,'').replace('-', '_') )
				$ovTask 					= 'oneview_appliance_device_snmp_{0}_users_facts' 		-f $_format
				$comment 					= '# ---------- Appliance {0} traps ' 					-f $format								
				$title 						= ' Get snmpV3 user {0} id' 							-f $snmpV3User
				[void]$YMLscriptCode.Add((Generate-ymlTask 		 		-title $title -comment $comment -OVTask $ovTask -isData $False))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix username	-value $snmpV3User	-indentlevel 2 ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix set_fact 		-isVar $True 										-indentlevel 1 ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix $var_user	-value "'{{appliance_device_snmp_v3_users[0].id}}'"		-indentlevel 2 ))
			}

			$ovTask 					= 'oneview_appliance_device_snmp_{0}_trap_destinations' -f $_format								
			$title 						= ' Create {0} trap' 									-f $format

			[void]$YMLscriptCode.Add((Generate-ymlTask 		 		-title $title  -OVTask $ovTask))
			
			if ($format -eq 'snmpV3')
			{
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -Prefix 	destinationAddress 	-value "'$destinationAddress'"	-indentlevel $indentDataStart ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -Prefix 	type 				-value Destination			-indentlevel $indentDataStart ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -Prefix 	userId 				-value "'{{$var_user}}'"	-indentlevel $indentDataStart ))
			}
			else 
			{
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -Prefix 	destination 		-value "'$destinationAddress'"	-indentlevel $indentDataStart ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -Prefix 	communityString 	-value $communityString			-indentlevel $indentDataStart ))
			}

			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-Prefix 	port 				-value "'$port'"				-indentlevel $indentDataStart ))
		}	
	
	}
	
		
	 # ---------- Generate script to file
	 YMLwriteToFile -code $YMLscriptCode -file $YMLfiles
	
}


 Function Import-remoteSupport([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
 {
	 $PSscriptCode             = [System.Collections.ArrayList]::new()
	 $cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	 connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	 
	 $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	 if ($List)
	 {
		 [void]$PSscriptCode.Add($codeComposer)
		 $enable 			= $List.enabled
		 $companyName 		= $List.companyName
		 $username 			= $List.insightOnlineUsername
		 $password 			= $List.insightOnlinePassword
		 $optimizeOptIn 	= $List.optimizeOptIn
	 
		 if ($enable -eq 'True')
		 {
			$enableParam 		=  ' -enable ' 
			$optParam 			= if ($optimizeOptIn -eq 'True')	{ ' -OptimizeOptIn ${0}' -f $optimizeOptIn} else {''}
				

			generate-credentialCode -username $username -password $password -PSscriptCode $PSscriptCode
			$credParam 			= ' -InsightOnlineUsername "{0}" -InsightOnlinePassword $securepassword ' -f $username
			[void]$PSscriptCode.Add( (Generate-PSCustomVarCode -prefix ('Set-OVRemoteSupport -enable -CompanyName "{0}" {1} {2}' -f $companyName, $credParam, $optParam ) ))
		 }
		 else 
		 {
			[void]$PSscriptCode.Add( (Generate-PSCustomVarCode -prefix 'Set-OVRemoteSupport -disable ' ))
		 }

	   [void]$PSscriptCode.Add('Disconnect-OVMgmt')
	 }
  # ---------- Generate script to file
  writeToFile -code $PSscriptCode -file $ps1Files

 }


 ####### TO BE COMPLETED
 Function Import-ligsnmp([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
 {
	 $PSscriptCode             = [System.Collections.ArrayList]::new()
	 $cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	 connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	 
	 $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	 if ($List)
	 {
		 [void]$PSscriptCode.Add($codeComposer)
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
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'write-host -foreground CYAN "----- Configuring OV snmp {0} "' -isVar $False ))
		newLine

		# ------------ snmp Read Community string
		if ($communityString)
		{
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'write-host -foreground CYAN "----- Importing snmp read Community string "'  -isVar $False ))
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('Set-OVSnmpReadCommunity -name "{0} "' -f $readCommunity)  -isVar $False ))
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
			
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'write-host -foreground CYAN "----- Importing snmp trap destination"'  -isVar $False ))
			for ($i=0; $i -lt $destArray.Count; $i++)
			{
				$destinationParam 	= ' '
			}

			newLine
		}





		endBlock
		# Skip creating because resource already exists
		[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix 'else'-isVar $False))
		startBlock
		[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $name + ' already exists.') -isVar $False -indentlevel 1 ))
		endBlock


	 }
	 
	if ($List)
	{
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	}
   # ---------- Generate script to file
   writeToFile -code $PSscriptCode -file $ps1Files
 }		 

 ####### TO BE COMPLETED


 Function Import-user([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
 {
	 $PSscriptCode             = [System.Collections.ArrayList]::new()
	 $cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	 connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	 
	 $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	 if ($List)
	 {
		 [void]$PSscriptCode.Add($codeComposer)
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
		 [void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating Remote Support contacts {0} "' -f $firstName) -isVar $False ))

		 ifBlock 		-condition ('if ($null -eq (Get-OVRemoteSupportContact | where email -eq "{0}" ))' -f $email)
		 [void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('new-OVRemoteSupportContact {0}{1}{2}{3}' -f $defaultParam, $nameParam,$languageParam, $notesParam) -isVar $False -indentlevle 1))		 
		 endIfBlock

		 #Skip creating because resource already exists
		 elseBlock
		 [void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $email + ' already exists.') -isVar $False -indentlevel 1 ))
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
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating local users {0} "' -f $name) -isVar $False ))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix username -Value ("'{0}'" -f $name)  ))

		ifBlock			-condition ('if ($null -eq (get-OVUser -Name  "{0}" ))' -f $name)  #HKD
		generate-credentialCode -component 'Users' -username $name -password $password -PSscriptCode $PSscriptCode  -indentlevel 1

		$fullNameParam 			= if ($fullName)		{ ' -FullName "{0}" ' 		-f $fullName }			else {''}
		$emailParam 			= if ($emailAddress)	{ ' -EmailAddress "{0}" ' 	-f $emailAddress }		else {''}
		$officeParam 			= if ($officePhone) 	{ ' -OfficePhone "{0}" ' 	-f $officePhone }		else {''}
		$mobileParam 			= if ($mobilePhone) 	{ ' -MobilePhone "{0}" ' 	-f $mobilePhone }		else {''}

		$rolesParam 			= $null
		if ($roles)
		{
			$roles 					= "@('" + $roles.Replace($SepChar,"'$comma'") + "')"
			$rolesParam 			= ' -Roles $roles'
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'roles' 	-Value $roles -indentlevel 1))
		}


		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix scopePermissions -Value '@()' -indentlevel 1  ))
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
				$value 		= 'Get-OVScope -name "{0}" -ErrorAction SilentlyContinue ' -f $scopeName  #HKD
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix $scopeIndex -value $value -indentlevel 1))

				ifBlock 	-condition ('if ($null -ne ${0})' -f $scopeIndex) 	-indentlevel 1
				$spIndex 	= "sp$Index"
				$value 		= '@{' + '{0};scope=${1}' -f $role,$scopeIndex + '}'
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix $spIndex -value $value -indentlevel 2))
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix scopePermissions -Value ('$scopePermissions + ${0}' -f $spIndex) -indentlevel 2  ))
				endIfBlock -indentlevel 1

				$Index++
			}

			
		}

		# Ensure there is no space after backtick (`)

		ifBlock			-condition 'if ($scopePermissions)' -indentlevel 1
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('New-OVUser -username "{0}" -password "{1}" `' -f $name, $password) -isVar $False -indentlevel 2))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0}{1}{2}{3}`' -f $fullNameParam, $emailParam ,$officeParam, $mobileParam) -isVar $False -indentlevel 4))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0} -ScopePermissions $scopePermissions' -f $rolesParam) -isVar $False -indentlevel 4))
		endIfBlock		-indentlevel 1

		elseBlock 		-indentlevel 1
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('New-OVUser -username "{0}" -password "{1}" `' -f $name, $password) -isVar $False -indentlevel 2))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0}{1}{2}{3}`' -f $fullNameParam, $emailParam ,$officeParam, $mobileParam) -isVar $False -indentlevel 4))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0}' -f $rolesParam) -isVar $False -indentlevel 4))		
		endElseBlock 	-indentlevel 1

		endIfBlock		-condition 'Get-OVUser...'

		# Skip creating because resource already exists
		elseBlock
		[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $name + ' already exists.') -isVar $False -indentlevel 1 ))
		endElseBlock

	 }
	 
	if ($List)
	{
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	}
   # ---------- Generate script to file
   writeToFile -code $PSscriptCode -file $ps1Files
 }


 Function Import-addressPool([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
 {
	 $PSscriptCode            = [System.Collections.ArrayList]::new()
	 $cSheet, $sheetName   	= $sheetName.Split($SepChar)       # composer

	 connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	 
	 $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	 if ($List)
	 {
		 [void]$PSscriptCode.Add($codeComposer)
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

		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating address pools {0} "' -f $name) -isVar $False ))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('# -------------- Attributes for address Pools "{0}"' -f $name) -isVar $False -indentlevel 1))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'proceed' 		-Value '$False' ) )
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'poolType' 		-Value ('"{0}"' 	-f $poolType) ))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'rangeType' 		-Value ('"{0}"' 	-f $rangeType) ))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'startAddress' 	-Value ('"{0}"' 	-f $startAddress) ))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'endAddress' 		-Value ('"{0}"' 	-f $endAddress) ))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'deleteGenerated' -Value ('${0}' 		-f $deleteGenerated) ))

		if ($poolType -like 'ip*')
		{
			
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'name' 		-Value ('"{0}"' 	-f $name) ))
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'networkId' 	-Value ('"{0}"' 	-f $networkId) ))
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'subnetmask' 	-Value ('"{0}"' 	-f $subnetmask) ))
			$gwParam 		= $null
			if ($gateway)
			{
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'gateway' 	-Value ('"{0}"' 	-f $gateway) ))
				$gwParam 	= '-Gateway $gateway '
			}
			$dnsParam 		= $null
			if ($dnsServers)
			{
				$value 		= "@('" + $dnsServers.replace($SepChar, "'$Comma'") + "')" 
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'dnsServers' 	-Value $value ))
				$dnsParam 	= ' -DNSServers $dnsServers '
			}
			$domainParam 	= $null
			if ($domain)
			{
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'domain' 	-Value ('"{0}"' 	-f $domain) ))
				$domainParam = ' -Domain $domain'
			}

		}

		if ($poolType -like "ip*")
		{
			ifBlock		-condition 'if ($poolType -like "ip*")' 
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'subnet' -Value (" Get-OVAddressPoolSubnet | where networkId -eq '{0}'" -f $networkId ) -indentlevel 1) )

			ifBlock		-condition 'if ($subnet -ne $null) ' -indentlevel 1
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'addressPool' -Value 'Get-OVAddressPoolRange | where subnetUri -match ($subnet.uri)' -indentlevel 2) )
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'proceed' -value ('$null -eq ($addressPool.startStopFragments.startAddress -ne $startAddress)')  -indentlevel 2 ) )
			endIfBlock  -condition '$subnet....'	-indentlevel 1
			
			elseBlock		-indentlevel 1 	# generate new subnet
			$value 			= 'new-OVAddressPoolSubnet -NetworkId $networkId -SubnetMask $subnetMask ' + $gwParam + $dnsParam + $domainParam
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'subnet' -value  $value -indentlevel 2 ) )
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'proceed' -Value '$True' -indentlevel 2 ) )
			endElseBlock	-indentlevel 1 
			endIfBlock	-condition '$poolType....' 
		}
		else 
		{
			$value = 'Get-OVAddressPoolRange| where {($_.name -eq $poolType) -and ($_.rangeCategory -eq $rangeType)} '
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'addressPool' -Value $value ) )
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'proceed' -value ('$null -eq ($addressPool | where startAddress -eq $startAddress)')  ) )
			
			#### Delete Generated range if asked

			$value = 'Get-OVAddressPoolRange| where {($_.name -eq $poolType) -and ($_.rangeCategory -eq "Generated")} '
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'poolToDelete' -Value $value ) )
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '$poolToDelete |  Remove-OVAddressPoolRange -confirm:$False' -isVar $False ) )

		}

		ifBlock			-condition  'if ($proceed)'

		$addressParam 		= ' -Start $startAddress -End $endAddress '
		

		if ($poolType -like "ip*")
		{
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('New-OVAddressPoolRange -IPSubnet $subnet -name "{0}" {1}' -f $name,  $addressParam ) -isVar $False -indentlevel 1))
		}
		else 
		{
			if ($rangeType -eq 'Generated') 
			{
				$addressParam 		= '' 
			}
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('New-OVAddressPoolRange -PoolType $poolType -RangeType $rangeType {0} ' -f  $addressParam ) -isVar $False -indentlevel 1))	
		}
		endIfBlock

		# Skip creating because resource already exists
		elseBlock
		[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $name + ' already exists.') -isVar $False -indentlevel 2 ))
		endElseBlock


		newLine
	 }
	 
	 if ($List)
	 {
		 [void]$PSscriptCode.Add('Disconnect-OVMgmt')
	 }

	# ---------- Generate script to file
	writeToFile -code $PSscriptCode -file $ps1Files
 }

 Function Import-YMLaddressPool_ipv4([string]$sheetName, [string]$WorkBook, [string]$YMLfiles )
{
     $YMLscriptCode         = [System.Collections.ArrayList]::new()
     $cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer
 
     [void]$YMLscriptCode.Add((Generate-ymlheader -title 'Configure IP v4 address pools'))
     
     $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	 
	 $List 					= $List | where poolType -eq 'ipV4'
	 foreach ($addr in $List)
	 {
		$name 				  	= $addr.name
		$poolType			  	= $addr.poolType
		$rangeCategory			= $addr.rangeCategory
		$deleteGenerated		= [Boolean]($addr.deleteGenerated)
		$startAddress 			= $addr.startAddress
		$endAddress 			= $addr.endAddress
		$networkId	 			= $addr.networkId
		$subnetmask 			= $addr.subnetmask
		$gateway 				= $addr.gateway
		$dnsServers 			= $addr.dnsServers
		$domain 				= $addr.domain

		$_subnetName 			="subnet_{0}" -f $networkId
		$comment				= '# ---------- IP v4 address pool {0} on subnet {1}' 	-f $name, $networkId
		$title 					= 'Ensure the ID Pools IPV4 Subnet exists'		
		[void]$YMLscriptCode.Add((Generate-ymlTask 			-title $title -comment $comment -OVTask 'oneview_id_pools_ipv4_subnet'))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		name					-value $_subnetName				-indentlevel $indentDataStart ))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		networkId				-value $networkId				-indentlevel $indentDataStart ))		
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		subnetmask				-value $subnetmask				-indentlevel $indentDataStart ))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		gateway					-value $gateway					-indentlevel $indentDataStart ))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		domain 					-value $domain 					-indentlevel $indentDataStart ))

		if ($dnsServers)
		{
			$dnsArr 			= $dnsServers.Split($SepChar)
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	dnsServers 												-indentlevel $indentDataStart ))
			foreach ($_dns in $dnsArr)
			{
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix $_dns		-IsVar $True -IsItem $True 					-indentlevel ($indentDataStart +2)  ))
			}

		}
		
		$var_subnet 		= "subnet_{0}_uri"		-f $name.Trim().Replace($Space, '')
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix set_fact 		-isVar $True 									-indentlevel 1 ))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix $var_subnet 	-value "`"{{id_pools_ipv4_subnet['uri'] }}`" "	-indentlevel 2 ))		
		
		# ---- id pools
		$title 					= 'Ensure the IPV4 Range {0} exists'	-f $name		
		[void]$YMLscriptCode.Add((Generate-ymlTask 			-title $title  -OVTask 'oneview_id_pools_ipv4_range'))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		name					-value $name					-indentlevel $indentDataStart ))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		subnetUri				-value "'{{$var_subnet}}'"		-indentlevel $indentDataStart ))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		type					-value Range					-indentlevel $indentDataStart ))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		rangeCategory			-value $rangeCategory			-indentlevel $indentDataStart ))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		startStopFragments										-indentlevel $indentDataStart ))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			'{'			-isVar $True -isItem $True				-indentlevel ($indentDataStart+1)))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 				startAddress 	-value "$startAddress ,"		-indentlevel ($indentDataStart+2)))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 				endAddress 		-value "$endAddress ," 			-indentlevel ($indentDataStart+2)))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 				fragmentType 	-value FREE 					-indentlevel ($indentDataStart+2)))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			'}'						-isItem $True				-indentlevel ($indentDataStart+1)))
		
	 }


	 # ---------- Generate script to file
     YMLwriteToFile -code $YMLscriptCode -file $YMLFiles
}


 # ---------- Ethernet networks
 Function Import-ethernetNetwork([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
 {
     $PSscriptCode             = [System.Collections.ArrayList]::new()
     $cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer
 
     connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
     
     $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
     
     foreach ($net in $List)
     {
         $name               = $net.name  
         $type 				 = $net.type       
         $vLANType           = $net.ethernetNetworkType
         $vLANID             = $net.vLanId
         $subnetID 			 = $net.subnetID
         $ipV6subnetID 		 = $net.ipV6subnetID
         $pBandwidth         = (1000 * $net.typicalBandwidth).ToString()
         $mBandwidth         = (1000 * $net.maximumBandwidth).ToString()
         $smartlink          = $net.SmartLink
         $private            = $net.PrivateNetwork
         $purpose            = $net.purpose
         $scopes             = $net.scopes
 
         [void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating ethernet networks {0} "' -f $name) -isVar $False ))
         [void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'net' -Value ("get-OVNetwork -name '{0}' -ErrorAction SilentlyContinue" -f $name) )) #HKD
 
         ifBlock 		-condition 'if ($Null -eq $net )' 
         [void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('# -------------- Attributes for Ethernet network "{0}"' -f $name) -isVar $False -indentlevel 1))
         [void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'name' -Value ('"{0}"' -f $name) -indentlevel 1))
         [void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'vLANType' -Value ('"{0}"' -f $vLANType) -indentlevel 1))
 
         $vLANIDparam = $vLANIDcode = $null
 
         # --- vLAN
         if ($vLANType -eq 'Tagged')
         { 
 
             if (($vLANID) -and ($vLANID -gt 0)) 
             {
                 $vLANIDparam = ' -VlanID $VLANID'
                 [void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'vLANid' -Value ('{0}' -f $vLANID) -indentlevel 1))
 
             }
 
         }                
 
         # --- Bandwidth
         $pBWparam = $pBWCode = $null
         $mBWparam = $mBWCode = $null
 
         if ($pBandwidth) 
         {
 
             $pBWparam = ' -TypicalBandwidth $pBandwidth'
             [void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'pBandwidth' -Value ('{0}' -f $pBandwidth) -indentlevel 1))
 
         }
 
         if ($mBandwidth) 
         {
 
             $mBWparam = ' -MaximumBandwidth $mBandwidth'
             [void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'mBandwidth' -Value ('{0}' -f $mBandwidth) -indentlevel 1))
 
         }
 
         # --- subnet
         $subnetCode     = $null
         $subnetIDparam  = ''
         $IPv6subnetCode = $IPv6subnetIDparam = $null
         $subnetArray 	= @()
         
         if ($subnetID)
         {
             
             [void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'ipV4subnet' -Value ("Get-OVAddressPoolSubnet -NetworkID `'{0}`'" -f $subnetID ) -indentlevel 1))
             $subnetArray += '$ipV4subnet'
         }
 
         if ($ipV6subnetID)
         {
             
             [void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'ipV6subnet' -Value ("Get-OVAddressPoolSubnet -NetworkID `'{0}`'" -f $ipV6subnetID ) -indentlevel 1))
             $subnetArray += '$ipV6subnet'
         }
 
         if ($subnetArray)
         {
                 $value 	= '@({0})' -f ($subnetArray -join $COMMA)
                 [void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'subnet' -Value $value -indentlevel 1) )
                 $subnetIDparam 	= ' -subnet $subnet'
         }
 
 
 
         [void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'PLAN' -Value ('${0}' -f $private) -indentlevel 1))
         [void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'smartLink' -Value ('${0}' -f $smartLink)-indentlevel 1))
         [void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'purpose' -Value ('"{0}"' -f $purpose)-indentlevel 1))
 
         [void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix 'New-OVNetwork -Name $name  -PrivateNetwork $PLAN -SmartLink $smartLink -VLANType $VLANType  -purpose $purpose `' -isVar $False -indentLevel 1)) 
         [void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('{0}{1}{2}{3}' -f $vLANIDparam, $pBWparam, $mBWparam, $subnetIDparam) -isVar $False -indentLevel 4))
         newLine # to end the command
         # --- Scopes
         if ($scopes)
         {
             newLine
             [void]$PSscriptCode.Add( (Generate-PSCustomVarCode -Prefix 'object' -Value 'get-OVNetwork | where name -eq $name' -indentlevel 1))
             generate-scopeCode -scopes $scopes -indentlevel 1
 
         }
 
         endIfBlock 
 
         # Skip creating because resource already exists
         elseBlock
         [void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $name + ' already exists.') -isVar $False -indentlevel 1 ))
         endElseBlock
 
 
         newLine
         
 
     }
 
     
     if ($List)
     {
         [void]$PSscriptCode.Add('Disconnect-OVMgmt')
     }
 
     # ---------- Generate script to file
     writeToFile -code $PSscriptCode -file $ps1Files
     
 }

 Function Import-YMLethernetNetwork([string]$sheetName, [string]$WorkBook, [string]$YMLfiles )
{
     $YMLscriptCode         = [System.Collections.ArrayList]::new()
     $cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer
 
     [void]$YMLscriptCode.Add((Generate-ymlheader -title 'Configure Ethernet networks'))
     
     $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
     
     foreach ($net in $List)
     {
         $name               = $net.name  
         $type 				 = $net.type       
         $vLANType           = $net.ethernetNetworkType
         $vLANID             = $net.vLanId
         $subnetID 			 = $net.subnetID
         $ipV6subnetID 		 = $net.ipV6subnetID
         $pBandwidth         = (1000 * $net.typicalBandwidth).ToString()
         $mBandwidth         = (1000 * $net.maximumBandwidth).ToString()
         $smartlink          = $net.SmartLink
         $private            = $net.PrivateNetwork
         $purpose            = $net.purpose
		 $scopes             = if ($net.scopes) {$net.scopes.Split('|')} else {$Null}
		 
		 newLine	-code $YMLscriptCode
		 $comment 			= '# ---------- Ethernet network  {0}' 	-f $name
		 [void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix $comment  				-isItem $True  					-indentlevel 1 ))

		 if ($subnetID)
		 {
			$title 			= 'get uri for subnet {0}' 	-f $subnetID
			$value 			= "'{0}'" 					-f $subnetID
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix name  		-value $title	-isVar $True 			-indentlevel 1 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix oneview_id_pools_ipv4_subnet_facts 					-indentlevel 1 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix config		-value "'{{config}}'"					-indentlevel 2 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix set_fact 				-isVar $True 				-indentlevel 1 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix networkId	-value $value							-indentlevel 2 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix var_subnets	-value "'{{id_pools_ipv4_subnets}}'"	-indentlevel 2 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix set_fact 				-isVar $True 				-indentlevel 1 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix var_uri		-value "'{{item.uri}}'"					-indentlevel 2 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix loop		-value "'{{var_subnets}}'"				-indentlevel 1 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix when		-value "item.networkId == networkId"	-indentlevel 1 ))
		 }



		 $_category 		= $YMLtype540Enum.item('ethernet')
		 $title 			= 'Create ethernet network {0}' 	-f $name		
		 [void]$YMLscriptCode.Add((Generate-ymlTask 		 -title $title  -OVTask 'oneview_ethernet_network'))	
		 [void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		type					-value $_category				-indentlevel $indentDataStart ))
		 [void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		name					-value $name					-indentlevel $indentDataStart ))
		 [void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		ethernetNetworkType		-value $vLANType				-indentlevel $indentDataStart ))
		 [void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		purpose					-value $purpose					-indentlevel $indentDataStart ))
		 [void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		smartLink				-value $smartLink				-indentlevel $indentDataStart ))
		 [void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		privateNetwork			-value $private					-indentlevel $indentDataStart ))
		 [void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		vlanId					-value $vlanID					-indentlevel $indentDataStart ))
		 if ($subnetID)
		 {
		 	[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	subnetUri			-value "'{{var_uri}}'"				-indentlevel $indentDataStart ))
		 }
		 [void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		bandwidth												-indentlevel $indentDataStart ))
		 [void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			typicalBandwidth		-value $pBandwidth			-indentlevel ($indentDataStart  + 2) ))
		 [void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			maximumBandwidth		-value $mBandwidth			-indentlevel ($indentDataStart  + 2) ))         
 
		 foreach ($_scope in $scopes)
		 {
			newLine	-code $YMLscriptCode
			$title 			= 'get ethernet network {0}' 	-f $name
			$_varname 		= "var_{0}"	-f ($name -replace " ", '_')
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix name  					-value $title	-isVar $True 		-indentlevel 2 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix oneview_ethernet_network_facts 								-indentlevel 2 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	config					-value "'{{config}}'"			-indentlevel 3 )) 
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	name					-value "'$name'"				-indentlevel 3 )) 
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix set_fact 				-isVar $True 						-indentlevel 2 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	$_varname			-value "'{{ethernet_networks.uri}}'" -indentlevel 3 ))

			newLine	-code $YMLscriptCode
			$title 			= 'Update the scope {0} with new resource {1}' 	-f $_scope, $_varname
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix name  					-value $title	-isVar $True 					-indentlevel 2 ))	
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix oneview_scope 															-indentlevel 2 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	config					-value "'{{config}}'"						-indentlevel 3 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	state					-value "resource_assignments_updated" 		-indentlevel 3 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'data'																-indentlevel 3 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		name					-value $_scope							-indentlevel $indentDataStart ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'resourceAssignments'											-indentlevel $indentDataStart ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			'addedResourceUris'											-indentlevel ($indentDataStart + 1) ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 				"'{{$_varname}}'"  		-isVar $True -isItem $True 		-indentlevel 6 ))	

		 }
	 }
	 

	 # ---------- Generate script to file
     YMLwriteToFile -code $YMLscriptCode -file $YMLFiles
}



# ---------- networks
Function Import-fcNetwork([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $PSscriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	
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

		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating FC/FCOE networks {0} "' -f $name) -isVar $False ))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'net' -Value ("get-OVNetwork -type {1} -name '{0}' -ErrorAction SilentlyContinue  " -f $name, $type) )) #HKD

		ifBlock			-condition 'if ($Null -eq $net )' 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('# -------------- Attributes for FC network "{0}"' -f $name) -isVar $False -indentlevel 1))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'name' -Value ('"{0}"' -f $name) -indentlevel 1))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'type' -Value ('"{0}"' -f $type)-indentlevel 1))

		# --- Bandwidth
		$pBWparam = $pBWCode = $null
		$mBWparam = $mBWCode = $null

		if ($PBandwidth) 
		{

			$pBWparam = ' -typicalBandwidth $pBandwidth'
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'pBandwidth' -Value ('{0}' -f $pBandwidth) -indentlevel 1))

		}

		if ($MBandwidth) 
		{

			$mBWparam = ' -maximumBandwidth $mBandwidth'
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'mBandwidth' -Value ('{0}' -f $mBandwidth) -indentlevel 1))

		}

		# --- ManagedSan
		$SANParam  = $Null
		if ($managedSan)
		{
			$SANparam   = ' -ManagedSAN $managedSAN' 

			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'SANname' -Value ('"{0}"' -f $SANname) -indentlevel 1))
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'managedSAN' -Value ('Get-OVManagedSAN -Name $SANname') -indentlevel 1))		
		}


		# ---- FC or FCOE network
		$FCParam 	 = $linkParam	 =  $autologinParam = $vLanIdParam = $null
		if ($type -eq 'fcoe')
		{
			if (($vLANID) -and ($vLANID -gt 0)) 
			{

				$vLanIdParam  =   ' -vLanID $vLanId'
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'vLanId' -Value ('{0}' -f $vLANID) -indentlevel 1))
			
			}

				
		}
		else # FC network
		{
			$FCparam          = ' -FabricType $fabricType'
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'fabricType' -Value ('"{0}"' -f $fabricType)  -indentlevel 1))
			if ($fabrictype -eq 'FabricAttach')
			{

				if ($autologinredistribution)
				{

					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'autologinredistribution' -Value ('${0}' -f $autologinredistribution)  -indentlevel 1))
					$autologinParam     = ' -AutoLoginRedistribution $autologinredistribution'

				}

				if ($linkStabilityTime) 
				{

					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'LinkStabilityTime' -Value ('{0}' -f $LinkStabilityTime) -indentlevel 1))
					$linkParam  = ' -LinkStabilityTime $LinkStabilityTime'

				}

				$FCparam              += $autologinParam + $linkParam
			}
		}

		# Note : when using backstick (`) make sure that theree is no space after. Otherwise it is considered as escape teh space char and NOT line continuator
		newLine
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('New-OVNetwork -Name $name -Type $Type `') -isVar $False -indentlevel 1) )
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0}{1}{2}{3}{4}' -f $pBWparam, $mBWparam, $FCparam, $vLANIDparam, $SANparam) -isVar $False -indentlevel 4) )
		newLine # to end the command

		# --- Scopes
		if ($scopes)
		{
			newLine
			[void]$PSscriptCode.Add( (Generate-PSCustomVarCode -Prefix 'object' -Value 'get-OVNetwork -name $name -ErrorAction SilentlyContinue' -isVar $False -indentlevel 1))
			generate-scopeCode -scopes $scopes -indentlevel 1

		}

		endIfBlock 

		# Skip creating because resource already exists
		elseBlock
		[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $name + ' already exists.') -isVar $False -indentlevel 1 ))
		endElseBlock

		newLine


	}

	if ($List)
	{
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	}
    # ---------- Generate script to file
    writeToFile -code $PSscriptCode -file $ps1Files

}

Function Import-YMLfcNetwork([string]$sheetName, [string]$WorkBook, [string]$YMLfiles )
{
     $YMLscriptCode         = [System.Collections.ArrayList]::new()
     $cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer
 
     [void]$YMLscriptCode.Add((Generate-ymlheader -title 'Configure Fibre Channel networks'))
     
     $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
     
     foreach ($net in $List)
     {
		$name                   = $net.name
		$type                   = $net.type
		$fabricType             = $net.fabricType
		$managedSan 			= $net.managedSan	
		$vlanId             	= $net.vLanId
		$pBandwidth             = 1000 * $net.typicalBandwidth
		$mBandwidth             = 1000 * $net.maximumBandwidth
		$autoLoginRedistribution = $net.autoLoginRedistribution
		$linkStabilityTime		 = $net.linkStabilityTime
		$scopes             	 = if ($net.scopes) {$net.scopes.Split('|')} else {$Null}
		
		$comment 				= '# ---------- FC or FCOE network  {0}' 	-f $name


		if ($vlanId) # fcoe network
		{
			$title 					= 'Create fcoe network {0}' 				-f $name		
			[void]$YMLscriptCode.Add((Generate-ymlTask 			-title $title -comment $comment -OVTask 'oneview_fcoe_network'))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		name					-value $name					-indentlevel $indentDataStart ))	
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		vlanId					-value $vlanId					-indentlevel $indentDataStart ))		
		}
		else # fc network
		{					

			$_category 				= $YMLtype540Enum.item('fcnetwork')
			$title 					= 'Create fc network {0}' 				-f $name		
			[void]$YMLscriptCode.Add((Generate-ymlTask 			-title $title -comment $comment -OVTask 'oneview_fc_network'))				
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		type					-value $_category				-indentlevel $indentDataStart ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		name					-value $name					-indentlevel $indentDataStart ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		fabricType				-value $fabricType				-indentlevel $indentDataStart ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		linkStabilityTime		-value $linkStabilityTime		-indentlevel $indentDataStart ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		autoLoginRedistribution	-value $autoLoginRedistribution	-indentlevel $indentDataStart ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		bandwidth												-indentlevel $indentDataStart ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			typicalBandwidth		-value $pBandwidth			-indentlevel ($indentDataStart  + 1) ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			maximumBandwidth		-value $mBandwidth			-indentlevel ($indentDataStart  + 1) ))         
		}
		 foreach ($_scope in $scopes)
		 {
			newLine	-code $YMLscriptCode
			$title 			= 'get fc or fcoe network {0}' 	-f $name
			$_oneview_facts = if ($vlanId) {'oneview_fcoe_network_facts'} else {'oneview_fc_network_facts'}
			$_varname 		= "var_{0}"	-f ($name -replace " ", '_')
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix name  					-value $title	-isVar $True 		-indentlevel 2 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix $_oneview_facts			 									-indentlevel 2 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	config					-value "'{{config}}'"			-indentlevel 3 )) 
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	name					-value "'$name'"				-indentlevel 3 )) 
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix set_fact 				-isVar $True 						-indentlevel 2 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	$_varname			-value "'{{ethernet_networks.uri}}'" -indentlevel 3 ))

			newLine	-code $YMLscriptCode
			$title 			= 'Update the scope {0} with new resource {1}' 	-f $_scope, $_varname
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix name  					-value $title	-isVar $True 					-indentlevel 2 ))	
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix oneview_scope 															-indentlevel 2 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	config					-value "'{{config}}'"						-indentlevel 3 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	state					-value "resource_assignments_updated" 		-indentlevel 3 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'data'																-indentlevel 3 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		name					-value $_scope							-indentlevel $indentDataStart ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'resourceAssignments'											-indentlevel $indentDataStart ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			'addedResourceUris'											-indentlevel ($indentDataStart + 1) ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 				"'{{$_varname}}'"  		-isVar $True -isItem $True 		-indentlevel 6 ))	

		 }
	 }
	 

	 # ---------- Generate script to file
     YMLwriteToFile -code $YMLscriptCode -file $YMLFiles
}



# ---------- Network Sets
Function Import-NetworkSet([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
	$PSscriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	
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

		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating networkset {0} "' -f $name) -isVar $False ))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'net' -Value ("get-OVNetworkSet -name '{0}' -ErrorAction SilentlyContinue" -f $name) )) #HKD

		ifBlock			-condition 'if ($Null -eq $net )'  
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('# -------------- Attributes for Network Set "{0}"' -f $name) -isVar $False -indentlevel 1))

		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'name' -Value ('"{0}"' -f $name) -indentlevel 1))
			
		$pBWparam = $pbWCode = $null
		$mBWparam = $mBWCode = $null

		if ($PBandwidth) 
		{

			$pBWparam = ' -TypicalBandwidth $pBandwidth'
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'pBandwidth' -Value ('{0}' -f $pBandwidth) -indentlevel 1))

		}
		
		if ($MBandwidth) 
		{

			$mBWparam = ' -MaximumBandwidth $mBandwidth'
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'mBandwidth' -Value ('{0}' -f $mBandwidth) -indentlevel 1))

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

			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix networks -value $value -indentlevel 1))
		}

		# --- native NEtwork
		$untaggedParam 		= $null

		if ($nativeNetwork)
		{
			$untaggedParam 		= ' -UntaggedNetwork $nativeNetwork'
			$value 				= '"{0}"' -f $nativeNetwork
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix untaggedNetwork -value $value -indentlevel 1))
		}

		newLine
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('New-OVNetworkSet -Name $name{0}{1}{2}{3}' -f $pBWparam, $mBWparam, $netParam, $untaggedParam) -isVar $False -indentlevel 1))		

		endIfBlock 

        # Skip creating because resource already exists
        elseBlock
        [void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $name + ' already exists.') -isVar $False -indentlevel 1 ))
		endElseBlock
		
		newLine
	}
	
	if ($List)
	{
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	}
    # ---------- Generate script to file
    writeToFile -code $PSscriptCode -file $ps1Files

}

Function Import-YMLnetworkSet([string]$sheetName, [string]$WorkBook, [string]$YMLfiles )
{
     $YMLscriptCode         = [System.Collections.ArrayList]::new()
     $cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer
 
     [void]$YMLscriptCode.Add((Generate-ymlheader -title 'Configure network sets'))
     
     $List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
     
     foreach ($ns in $List)
     {
		$name               = $ns.name
		$pBandwidth         = 1000 * $ns.TypicalBandwidth 
		$mBandwidth         = 1000 * $ns.MaximumBandwidth 
		$networkSetType 	= $ns.networkSetType
		$networks        	= if ($ns.networks) {$ns.networks.Split('|')} else {$Null}
		$nativeNetwork 		= $ns.nativeNetwork
		$scopes             = if ($ns.scopes) {$ns.scopes.Split('|')} else {$Null}
		 
		 $comment 			= '# ---------- Network set  {0}' 	-f $name
		 $title 			= 'Create network set {0}' 			-f $name		
		 [void]$YMLscriptCode.Add((Generate-ymlTask 		-title $title -comment $comment -OVTask 'oneview_network_set'))	

		 [void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		name					-value $name					-indentlevel $indentDataStart ))
		 [void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		networkSetType			-value $networkSetType			-indentlevel $indentDataStart ))
		 [void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		bandwidth												-indentlevel $indentDataStart ))
		 [void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			typicalBandwidth	-value $pBandwidth				-indentlevel ($indentDataStart  + 1) ))
		 [void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			maximumBandwidth	-value $mBandwidth				-indentlevel ($indentDataStart  + 1) ))         

		 if ($nativeNetwork)
		 {
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		nativeNetworkUri	-value $nativeNetwork			-indentlevel $indentDataStart ))
		 }
		 if ($networks)
		 {
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		networkUris											-indentlevel $indentDataStart )) 
			foreach ($_net in $networks)
			{
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		$_net		-isVar $True -isItem $True			-indentlevel ($indentDataStart+1) )) 
			}
		 }

		 foreach ($_scope in $scopes)
		 {
			newLine	-code $YMLscriptCode
			$title 			= 'get fc network {0}' 	-f $name
			$_varname 		= "var_{0}"	-f ($name -replace " ", '_')
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix name  					-value $title	-isVar $True 		-indentlevel 2 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix oneview_network_set_facts 									-indentlevel 2 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	config					-value "'{{config}}'"			-indentlevel 3 )) 
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	name					-value "'$name'"				-indentlevel 3 )) 
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix set_fact 				-isVar $True 						-indentlevel 2 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	$_varname			-value "'{{ethernet_networks.uri}}'" -indentlevel 3 ))

			newLine	-code $YMLscriptCode
			$title 			= 'Update the scope {0} with new resource {1}' 	-f $_scope, $_varname
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix name  					-value $title	-isVar $True 					-indentlevel 2 ))	
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix oneview_scope 															-indentlevel 2 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	config					-value "'{{config}}'"						-indentlevel 3 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	state					-value "resource_assignments_updated" 		-indentlevel 3 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'data'																-indentlevel 3 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		name					-value $_scope							-indentlevel $indentDataStart ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'resourceAssignments'											-indentlevel $indentDataStart ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			'addedResourceUris'											-indentlevel ($indentDataStart + 1) ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 				"'{{$_varname}}'"  		-isVar $True -isItem $True 		-indentlevel 6 ))	

		 }
	 }
	 

	 # ---------- Generate script to file
     YMLwriteToFile -code $YMLscriptCode -file $YMLFiles
}

# ---------- LIG
Function Import-LogicalInterconnectGroup([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
	$PSscriptCode             	= [System.Collections.ArrayList]::new()

	$cSheet, $ligSheet, $uplSheet, $snmpConfigSheet, $snmpV3UserSheet, $snmpTrapSheet 	= $sheetName.Split($SepChar)       # composer
	
	$ligList 					= if ($ligSheet)	 		{get-datafromSheet -sheetName $ligSheet -workbook $WorkBook				} else {$null}
	
	#Note: snmp____List will be extracted in teh snmp subsection

	

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	
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


		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating logical interconnect group {0} " ' -f $name) -isVar $False ))

		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'lig' 						-Value ("get-OVLogicalInterconnectGroup -name  '{0}' -ErrorAction SilentlyContinue" -f $name) )) #HKD

		ifBlock			-condition 'if ($lig -eq $Null)' 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('# -------------- Attributes for LIG "{0}"' -f $name) -isVar $False -indentlevel 1))

		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'name' 						-Value ('"{0}"'	-f $name) -indentlevel 1))
		
		# --- Frame Count - InterconnectBay Set - Fabric Module Type
		$FrameCountParam		= ' -frameCount $frameCount'
		$ICBaySetParam			= ' -interConnectBaySet $interConnectBaySet '
		$fabricModuleParam 		= ' -fabricModuleType $fabricModuleType'
		
		
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'frameCount' 					-Value ('{0}' 	-f $frameCount) -indentlevel 1))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'interconnectBaySet'			-Value ('{0}' 	-f $ICBaySet) -indentlevel 1))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'fabricModuleType' 			-Value ("'{0}'" -f  $fabricModuleType) -indentlevel 1 )) 

		# redundancy Type
		$redundancyParam 	= $null
		if ($redundancyType)
		{
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'redundancyType' 				-Value ('"{0}"'	-f $redundancyType) -indentlevel 1))
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

			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'bayConfig' -Value  $value -indentlevel 1 ))

			$baysParam 		= ' -Bays $bayConfig'

			

		}

		if ($isNotSAS)
		{
			$igmpParam = $igmpIdleTimeoutParam = $null
			if ($enableIgmpSnooping -eq 'True')
			{
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'enableIgmpSnooping'			-Value ('${0}' 	-f $enableIgmpSnooping) -indentlevel 1))
				if ($igmpIdleTimeoutInterval)
				{
					$igmpIdleTimeoutParam = ' -igmpIdleTimeOutInterval $igmpIdleTimeoutInterval'
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'igmpIdleTimeoutInterval'		-Value ('{0}' 	-f $igmpIdleTimeoutInterval) -indentlevel 1))
				}
				$igmpParam                  = ' -enableIgmpSnooping $enableIgmpSnooping {0}' -f $igmpIdleTimeoutParam
			}


			
			$networkLoopProtectionParam 	= ' -enablenetworkLoopProtection $enableNetworkLoopProtection'
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'enableNetworkLoopProtection'	-Value ('${0}' 	-f $enableNetworkLoopProtection) -indentlevel 1))
			

			$EnhancedLLDPTLVParam       	= ' -enableEnhancedLLDPTLV $enableRichTLV'
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'enableRichTLV'				-Value ('${0}' 	-f $enableRichTLV) -indentlevel 1))

			$LLDPtaggingParam 		      	= ' -enableLLDPTagging $enableTaggedLldp'
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'enableTaggedLldp'			-Value ('${0}' 	-f $enableTaggedLldp) -indentlevel 1))

			$LldpAddressingModeParam		= ' -lldpAddressingMode $lldpIpAddressMode'
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'lldpIpAddressMode'			-Value ("'{0}'" -f $lldpIpAddressMode) -indentlevel 1))			

			#$stormControlParam 				= $null
			#if ($enableStormControl -eq 'True')
			#{
			#	$stormControlParam 				= ' -enableStormControl $enableStormControl '
			#	[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'enableStormControl'			-Value ('${0}' 	-f $enableStormControl) -indentlevel 1))
			#	[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'stormControlPollingInterval'	-Value ('{0}' 	-f $stormControlPollingInterval) -indentlevel 1))
			#	[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'stormControlThreshold'		-Value ('{0}' 	-f $stormControlThreshold) -indentlevel 1))
			#}

			# --- Specific to C7000
			
			if ($enclosureType -ne	$Syn12K)
			{
				$macCacheParam 					= $null
				if ($enableFastMacCacheFailover -eq 'True')
				{
					$macCacheParam 				= ' -enableFastMacCacheFailover $enableFastMacCacheFailover -MacRefreshInterval $macRefreshInterval'
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'enableFastMacCacheFailover'	-Value ('${0}' 	-f $enableFastMacCacheFailover) -indentlevel 1))
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'macRefreshInterval'			-Value ('{0}' 	-f $macRefreshInterval) -indentlevel 1))
				}
				$pauseFloodProtectionParam		= $null
				$pauseFloodProtectionParam 		= ' -enablePauseFloodProtection $enablePauseFloodProtection'
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'enablePauseFloodProtection'	-Value ('${0}' 	-f $enablePauseFloodProtection) -indentlevel 1))

			}

			$InterconnectConsistencyCheckingParam = ' -interconnectConsistencyChecking $interconnectConsistency'
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'interconnectConsistency'		-Value ("'{0}'" 	-f $interconnectConsistencyChecking) -indentlevel 1))

			# ------ Internal Networks
			$intNetParam 	= $null
			if ($internalNetworks)
			{

				$networks 	= $internalNetworks.replace($sepChar, '";"')
				$networks 	= $networks.Insert($networks.length,'")').Insert(0,'@("')
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'internalNetworks'			-Value ('{0}' -f $networks + ' | % {Get-OVNetwork -name $_}' 	) -indentlevel 1))
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'internalNetworkConsistency'	-Value ("'{0}'" -f $internalNetworkConsistency ) -indentlevel 1))				
			
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
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'communityString'	-Value ("'{0}'" -f $communityString ) -indentlevel 1))
					$communityStringParam 	= ' -snmpV1 $True -ReadCommunity $communityString '
				}

				if ($contact) 
				{
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'contact'	-Value ("'{0}'" -f $contact ) -indentlevel 1))
					$contactParam 	= ' -Contact $contact '					
				}
				if ($accList)
				{
					$accList 			= "@(" + ($accList -replace '|', ',') + ")"
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'accessList'	-Value ("{0}" -f $accList ) -indentlevel 1))	
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
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'snmpConfiguration' -value ('New-OVSnmpConfiguration {0}{1}{2}{3}{4} ' -f $communityStringParam, $contactParam, $accessListParam, $snmpV3userParam, $snmpTrapDestinationParam)  -indentlevel 1))
					newLine
					$snmpConfigurationParam 	= ' -snmp $SnmpConfiguration '
				}


			}

			

			$ligVariable    = '$lig'
			newLine
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0} = New-OVLogicalInterconnectGroup -Name $name {1}{2}{3} `' -f $LigVariable, $fabricModuleParam, $FrameCountParam , $ICBaySetParam) -isVar $false -indentlevel 1))
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0}{1} `' 		-f $baysParam, $redundancyParam ) -isVar $false -indentlevel 4))
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0}{1}{2}{3} `' -f $intNetParam,$igmpParam,$pauseFloodProtectionParam, $networkLoopProtectionParam) -isVar $false -indentlevel 4))
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0}{1} `' 		-f $macCacheParam,$EnhancedLLDPTLVParam ) -isVar $false -indentlevel 4))
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0}{1}{2} `' 	-f $LldpAddressingModeParam,$LLDPtaggingParam,$InterconnectConsistencyCheckingParam ) -isVar $false -indentlevel 4))
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0} '			-f $snmpConfigurationParam ) -isVar $false -indentlevel 4))
			
			newLine # to end the command

			#TBD      ,  $snmpParam, $QosParam, $ScopeParam))

		}
		else  # SAS lig
		{
			$ligVariable    = '$lig'
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0} = New-OVLogicalInterconnectGroup -Name $name {1}{2}{3}{4}' -f $LigVariable, $fabricModuleParam, $FrameCountParam , $ICBaySetParam, $baysParam) -isVar $false -indentlevel 1))
				
		}

		endIfBlock -condition 'if ($lig -eq $Null)' 

		# Skip creating because resource already exists
		elseBlock
		[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW ' + "'{0}'" -f $name + ' already exists.') -isVar $False -indentlevel 1 ))
		endElseBlock

		newLine

	}
	
	if ($List)
	{
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	}

	 # ---------- Generate script to file
	 writeToFile -code $PSscriptCode -file $ps1Files
}

# ---------- Uplink Set
Function Import-UplinkSet([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $PSscriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	
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
		
		
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating uplinkset {0} on LIG {1}"' -f $uplName,$ligName) -isVar $False ))
 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'lig' 						-Value ("get-OVLogicalInterconnectGroup -name '{0}' -ErrorAction SilentlyContinue" -f $ligName ) ))	#HKD	
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'upl' 						-Value ('$lig.uplinksets | where name -eq  "{0}" ' -f $uplName) ))
		newLine

		ifBlock			-condition 'if ( ($lig -ne $Null) -and ($upl -eq $Null) )' 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('# -------------- Attributes for uplinkset {0} on LIG {1}' -f $uplName,$ligName) -isVar $False -indentlevel 1))

		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'name' 						-Value ('"{0}"'	-f $uplName) -indentlevel 1))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'networkType' 				-Value ('"{0}"'	-f $networkType) -indentlevel 1))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'uplConsistency' 				-Value ('"{0}"'	-f $consistency) -indentlevel 1))


		# ---- Networks
		$netParam = $null
		if ($networks)
		{
			$netParam   = ' -Networks $networks'

			$arr 		= [system.Collections.ArrayList]::new()
			$arr 		= $networks.split($SepChar) | % { '"{0}"' -f $_ }
			$netList    = '@({0})' -f [string]::Join($Comma, $arr)
			$value 		= ('{0}' -f $netList ) + ' | % { get-OVNetwork -name $_ }'

			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix networks -value $value -indentlevel 1))
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
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix networkSets -value $value -indentlevel 1))
			}

			# ----- Nativenetwork
			if ($nativeNetwork)
			{
				$nativeNetParam = ' -NativeEthNetwork $nativeNetwork'
				$value 			= "get-OVNetwork -name '{0}' " -f $nativeNetwork
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix nativeNetwork -value $value -indentlevel 1))
			}

			# ---- lacpTimer and loadbalancing
			$lacpParam 	= ' -LacpTimer $lacpTimer -LacpLoadbalancingMode $lacpLoadbalancingMode'
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix lacpTimer -value ("'{0}'" -f $lacpTimer) -indentlevel 1))
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix lacpLoadbalancingMode -value ("'{0}'" -f $loadbalancingMode) -indentlevel 1))
		}
		else # Fibre Channel specific
		{
			$trunkingParam 	= ' -enableTrunking $enableTrunking'
			$value 			= '${0}' -f $enableTrunking
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix enableTrunking -value $value -indentlevel 1))

			$trunkingParam 	= ' -fcUplinkSpeed $fcUplinkSpeed'
			$value 			= "'{0}'" -f ($fcSpeed -replace 'Gb','')
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix fcUplinkSpeed  -value $value -indentlevel 1))
		}

		# ---- Logical Ports config -- transform Enclosure1:Bay3:Q1|Enclosure1:Bay3:Q2|Enclosure1:Bay3:Q3 into table
		$uplinkPortParam 	= $null
		if ($logicalPortConfigInfos)
		{
			$uplinkPortParam 	= ' -UplinkPorts $uplinkPorts'
			$value 				= $logicalPortConfigInfos.replace($SepChar, '","') 	# Comma and quote
			$value 				= "@(`"$value`")"
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix uplinkPorts -value $value -indentlevel 1))
		}




		# Make sure there is no space after backtick (`)
		newLine
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -prefix ('New-OVUplinkSet -InputObject $lig -Name $name -Type $networkType `') -isVar $False -indentlevel 1)) 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -prefix ('{0}{1}{2}{3}{4} `' 	-f $netParam, $netsetParam, $nativeNetParam, $trunkingParam, $lacpParam) -isVar $False -indentlevel 4)) 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -prefix ('{0} `' 				-f $uplinkPortParam) -isVar $False -indentlevel 4)) 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -prefix (' -ConsistencyChecking $uplConsistency' ) -isVar $False -indentlevel 4))
		newLine # to end the command
		
		endIfBlock

		# Skip creating because resource already exists
		elseBlock
		[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW ' + "{0} does not exist or " -f $ligName   + "{0} already exists." -f $uplName ) -isVar $False -indentlevel 1 ))
		endElseBlock

		newLine
        
        
	}
	
	if ($List)
	{
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	}
	
	 # ---------- Generate script to file
	 writeToFile -code $PSscriptCode -file $ps1Files



}

Function Import-YMLUplinkSet([string]$sheetName, [string]$WorkBook, [string]$YMLfiles )
{
    $YMLscriptCode          = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	[void]$YMLscriptCode.Add((Generate-ymlheader -title 'Configure uplink set'))
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  	

	$definedVarPort          = [System.Collections.ArrayList]::new()
	$definedVarNet 	         = [System.Collections.ArrayList]::new()


	foreach ( $upl in $List)
	{
		$ligName                    = $upl.ligName
        $uplName           			= $upl.name
		$networkType 	            = $upl.networkType
		$networks 			        = $upl.Networks			#[]
		$networkSets				= $upl.NetworkSets		#[]
		$nativeNetwork 		    	= $upl.nativeNetwork
		$enableTrunking 		    = $upl.enableTrunking
		$fabricModuleName			= $upl.fabricModuleName
		$logicalPortConfigInfos		= $upl.LogicalPortConfigInfos

		$lacpTimer       			= if ($upl.lacpTimer) 			{  $upl.lacpTimer.Trim() } 			else { 'Short' }
		$loadBalancingMode			= $upl.loadBalancingMode
		$primaryPort     			= $upl.PrimaryPort
		$fcSpeed         			= $upl.FCuplinkSpeed #[]
		$consistency 				= if ($upl.consistencyChecking) {$YMLconsistencyCheckingLigEnum.Item($upl.consistencyChecking) } else {'None'}
		
		newLine	-code $YMLscriptCode
		$comment 			= '# ---------- Uplink set {0} for LIG {1}' 	-f $uplName, $ligName
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix $comment  				-isItem $True  					-indentlevel 1 ))

		
		#### --- Query section 
		#### --------- Set variables for various attributes:
		####   --- Port name on Interconnect Type
		####   --- URi for FC network
		####   --- URi for native network
		####   --- URi for network set


		#---- Port name for Interconnect
		if ($logicalPortConfigInfos)
		{
			$ligNameStr 			= $ligName.replace(' ', '_').replace('-', '_')   	#build ligname for variable
			$ICNameStr 				= $ICModuleTypes.item($fabricModuleName.replace($Space,'').Trim())

			$portInfoArr 			= $logicalPortConfigInfos.Split($SepChar)
			foreach ($_p in $portInfoArr)
			{
				if ($fabricModuleName -like '*FC*')						# If FibreChannel module
				{
					$bay,$port		= $_p.Trim().Split(':')    	# Bay5:7|Bay5:8
					# Start with number
					$port = if ($port -match "^[0-9]" ) {"Q$port"} else {$port} 						# Like Q1, Q2
				}
				else 
				{
					$encl,$bay,$port	= $_p.Trim().Split(':')	
					$_frameNumber 		= $encl[-1]
				}

				$_bayNumber			= $bay[-1]
				if ($port)		# works only with Synergy not C7000 syntax Bay 2:2
				{
					$_portName 		= $port.replace('.', ':') 	# normalize port naming convention
					$portStr 		= $port.replace('.', '_')
				}

				$varPort 			= "var_{0}_{1}"		-f $ICNameStr , $portStr   # var_VC40F8_Q1_1
				if ($definedVarPort -notcontains $varPort  )		# Check whether it's the same port name Qxxx 
				{
					# Now we get portName like Q7:1 and IC name - build Ansible task to collect port Number
					newLine	-code $YMLscriptCode
					$_task 				= ' Query OneView to get interconnect types and port number from lig {0} and port name {1} ' 		-f $ligName, $port			
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix name  				-value $_task	-isVar $True 			-indentlevel 1 ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix oneview_interconnect_type_facts								-indentlevel 1 ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	config			-value "'{{config}}'"					-indentlevel 2 ))

					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix set_fact			-isVar $True 							-indentlevel 1 ))	
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	list_portInfos	-value "'{{item.portInfos}}'"			-indentlevel 2 ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix loop				-value "'{{interconnect_types}}'"		-indentlevel 1 ))	
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix when				-value "item.name =='$fabricModuleName'" -indentlevel 1 ))

					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix set_fact			-isVar $True 							-indentlevel 1 ))	
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix $varPort 			-value "'{{item.portNumber}}'" 			-indentlevel 2 ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix loop				-value "'{{list_portInfos}}'"			-indentlevel 1 ))	
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix when				-value "item.portName =='$_portName'" 	-indentlevel 1 ))

					[void]$definedVarPort.Add($varPort)
				}
			}

		}

		#---- URI for FC network
		if ($networkType -eq 'Ethernet' )
		{
			if ($nativeNetwork)
			{
				$_netStr 		= $nativeNetwork.replace($Space,'_').replace('-', '_').Trim()
				$var 			= "var_{0}"	-f $_netStr
			
				if ($definedVarNet -notcontains $var  )		# Check whether it's the same port name Qxxx 
				{
					newLine	-code $YMLscriptCode
					$_task 			= ' get URI for network {0} ' -f $_net 				
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix name  				-value $_task	-isVar $True 			-indentlevel 1 ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix oneview_ethernet_network_facts								-indentlevel 1 ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	config			-value "'{{config}}'"					-indentlevel 2 ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	name			-value $_net							-indentlevel 2 ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix set_fact							-isVar $True 			-indentlevel 1 ))	
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	$var			-value "'{{ethernet_networks[0].uri}}'"	-indentlevel 2 ))	
					
					[void]$definedVarPort.Add($var)
				}	
			}

			if ($networkSets)
			{
				$netArr 			= $networkSets.Split($SepChar)
				foreach ($_net in $netArr)
				{
					$_netStr 		= $_net.replace($Space,'_').replace('-', '_').Trim()
					$var 			= "var_network_set_{0}"	-f $_netStr

					if ($definedVarNet -notcontains $var  )		# Check whether it's the same port name Qxxx 
					{
						newLine	-code $YMLscriptCode
						$_task 			= ' get URI for network set {0} ' -f $_net 				
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix name  				-value $_task	-isVar $True 			-indentlevel 1 ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix oneview_network_set_facts									-indentlevel 1 ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	config			-value "'{{config}}'"					-indentlevel 2 ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	name			-value $_net							-indentlevel 2 ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix set_fact							-isVar $True 			-indentlevel 1 ))	
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	$var			-value "'{{network_sets[0].uri}}'"		-indentlevel 2 ))	

						[void]$definedVarPort.Add($var)
					}
				}
			}
		}
		else # Fibre Channel
		{
			if ($networks)
			{
				$netArr 			= $networks.Split($SepChar)
				foreach ($_net in $netArr)
				{
					$_netStr 		= $_net.replace($Space,'_').replace('-', '_').Trim()
					$var 			= "var_{0}"	-f $_netStr
				
					if ($definedVarNet -notcontains $var  )		# Check whether it's the same port name Qxxx 
					{
						newLine	-code $YMLscriptCode
						$_task 			= ' get URI for network {0} ' -f $_net 				
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix name  				-value $_task	-isVar $True 			-indentlevel 1 ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix oneview_fc_network_facts									-indentlevel 1 ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	config			-value "'{{config}}'"					-indentlevel 2 ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	name			-value $_net							-indentlevel 2 ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix set_fact							-isVar $True 			-indentlevel 1 ))	
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	$var			-value "'{{fc_networks[0].uri}}'"			-indentlevel 2 ))	

						[void]$definedVarPort.Add($var)
					}
				}

			}
		}

		#### --- End Query section 

		$title 			= ' Create uplink set {0} for LIG {1}' 			-f $uplName, $ligName	
		[void]$YMLscriptCode.Add((Generate-ymlTask 		 -title $title -OVTask 'oneview_logical_interconnect_group'))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'name'				-value $ligName							-indentlevel $indentDataStart ))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'uplinkSets'												-indentlevel $indentDataStart ))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			'name'			-value $uplName 		-isVar $True	-indentlevel ($indentDataStart+1) ))
		
		# ---- Networks
		if ($networks)
		{
			if ($networkType  -eq 'Ethernet')
			{
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'networkNames' 											-indentlevel ($indentDataStart+1) ))
				$netArr					= $networks.Split($SepChar)
				foreach ($net in $netArr)
				{
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 				"- $net"  				-isItem	 $True		-indentlevel ($indentDataStart+2) ))
				}
			}
			else # Fiber Channel
			{
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'networkUris' 											-indentlevel ($indentDataStart+1) ))
				$netArr					= $networks.Split($SepChar)
				foreach ($net in $netArr)
				{
					$_netStr 		= $net.replace($Space,'_').replace('-', '_').Trim()
					$var 			= "var_{0}"	-f $_netStr
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		"- '{{$var}}'"  					-isItem	 $True		-indentlevel ($indentDataStart+2) ))
				
				}

			}
		}

		# ---- logical ports
		if ($logicalPortConfigInfos)
		{
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'logicalPortConfigInfos'									-indentlevel ($indentDataStart+1) ))
			
			$ICNameStr 				= $ICModuleTypes.item($fabricModuleName.replace($Space,'').Trim())
			$portInfoArr 			= $logicalPortConfigInfos.Split($SepChar)
			foreach ($_p in $portInfoArr)
			{
				if ($ligName -like '*FC*')						# If FibreChannel module
				{
					$bay,$port		= $_p.Trim().Split(':')    	# Bay5:7
					# Start with number
					$port = if ($port -match "^[0-9]" ) {"Q$port"} else {$port} 						# Like Q1, Q2

				}
				else 
				{
					$encl,$bay,$port	= $_p.Trim().Split(':')	
				}
				$_frameNumber 		= if ($ligName -like '*FC*') {-1} else {$encl[-1]}
				$_bayNumber			= $bay[-1]
				if ($port)		# works only with Synergy not C7000 syntax Bay 2:2
				{
					$_portName 		= $port.replace('.', ':') 	# normalize port naming convention
					$portStr 		= $port.replace('.', '_')
				}
				$varPort 			= "var_{0}_{1}"		-f $ICNameStr , $portStr   # var_VC40F8_Q1_1
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix desiredSpeed 	-Value Auto	-isVar $True 						-indentlevel ($indentDataStart+3) ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	logicalLocation												-indentlevel ($indentDataStart+3) ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		locationEntries											-indentlevel ($indentDataStart+4) ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			type			-Value 'Enclosure' 	-isVar $True	-indentlevel ($indentDataStart+5) ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			relativeValue	-Value $_frameNumber 				-indentlevel ($indentDataStart+5) ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			type			-Value 'Bay' 		-isVar $True	-indentlevel ($indentDataStart+5) ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			relativeValue	-Value $_bayNumber 					-indentlevel ($indentDataStart+5) ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			type			-Value 'Port' 		-isVar $True	-indentlevel ($indentDataStart+5) ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			relativeValue	-Value "'{{$varPort}}'"  			-indentlevel ($indentDataStart+5) ))
			}

		}


		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			'networkType'			-value $networkType 				-indentlevel ($indentDataStart+1) ))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			'mode'					-value Auto 						-indentlevel ($indentDataStart+1) ))
		if ($networkType -eq 'Ethernet') 
		{

			# ---- lacpTimer and loadbalancing and connection Mode
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'lacpTimer'				-value $lacpTimer					-indentlevel ($indentDataStart+1) ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'loadBalancingMode'		-value $loadBalancingMode			-indentlevel ($indentDataStart+1) ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'consistencyChecking'	-value $consistency 				-indentlevel ($indentDataStart+1) ))

			# ---- Network Sets ( No CopyNetworksFromNetworkSet)
			if ($networkSets)
			{
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'networkSetUris' 											-indentlevel ($indentDataStart+1) ))
				$netArr					= $networkSets.Split($SepChar)
				foreach ($net in $netArr)
				{
					$_netStr 		= $net.replace($Space,'_').replace('-', '_').Trim()
					$var 			= "var_{0}"	-f $_netStr
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		"- '{{$var}}'"  				-isItem	 $True		-indentlevel ($indentDataStart+2) ))
				
				}
				
			}

			# ----- Nativenetwork
			if ($nativeNetwork)
			{
				$_netStr 		= $nativeNetwork.replace($Space,'_').replace('-', '_').Trim()
				$var 			= "var_{0}"	-f $_netStr			
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'nativeNetworkUri' 	-value "'{{$var}}'" 					-indentlevel ($indentDataStart+1) ))
				
			}

		}
		else # Fibre Channel specific
		{
			$trunkingParam 	= ' -enableTrunking $enableTrunking'
			$value 			= '${0}' -f $enableTrunking
			

			$trunkingParam 	= ' -fcUplinkSpeed $fcUplinkSpeed'
			$value 			= "'{0}'" -f ($fcSpeed -replace 'Gb','')
			
		}


        
        
	}
	

	
	 # ---------- Generate script to file
	 YMLwriteToFile -code $YMLscriptCode -file $YMLfiles



}

Function Import-YMLligSNMP([string]$sheetName, [string]$WorkBook, [string]$YMLfiles )
{
    $YMLscriptCode          										= [System.Collections.ArrayList]::new()
	$cSheet, $snmpConfigSheet, $snmpV3UserSheet, $snmpTrapSheet 	= $sheetName.Split($SepChar)       # composer

	[void]$YMLscriptCode.Add((Generate-ymlheader -title 'Configure SNMP for LIG'))
	
	$snmpConfigList 		= get-datafromSheet -sheetName $snmpConfigSheet -workbook $WorkBook  	
	$snmpConfigList			= $snmpConfigList | where source -ne 'Appliance'

	$snmpV3UserList 		= get-datafromSheet -sheetName $snmpV3UserSheet -workbook $WorkBook  	
	$snmpV3UserList			= $snmpV3UserList | where source -ne 'Appliance'

	$snmpTrapList 			= get-datafromSheet -sheetName $snmpTrapSheet -workbook $WorkBook  	
	$snmpTrapList			= $snmpTrapList | where source -ne 'Appliance'


	foreach ( $_snmp in $snmpConfigList)
	{	
		# snmpV1 information
		$ligName 				= $_snmp.source
		$consistencyChecking	= $_snmp.consistencyChecking
		$communityString		= $_snmp.communityString
		$contact 				= $_snmp.contact
		$accList 				= $_snmp.accessList

		$comment 			= '# ---------- snmp Configuration for LIG {0}' 				-f $ligName
		$title 				= ' Create snmp Configuration for LIG {0}' 						-f $ligName	

		[void]$YMLscriptCode.Add((Generate-ymlTask 		 -title $title -comment $comment -OVTask 'oneview_logical_interconnect_group'))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'name'				-value $ligName					-indentlevel $indentDataStart ))
	
		if ($communityString)
		{
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	readCommunity		-Value $communityString  				-indentlevel $indentDataStart ))
		}
	
		if ($contact) 
		{
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -Prefix 	systemContact		-Value $contact  						-indentlevel $indentDataStart ))		
		}

		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -Prefix 	v3Enabled 					-value 'True'						-indentlevel $indentDataStart ))
	
		# ----  snmpV3Users for this lig
		$_snmpV3UserList				= $snmpV3UserList | where source -eq $ligName
		if ($_snmpV3UserList)
		{
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode	-Prefix 	'snmpUsers: [ ' 			-isItem $True					-indentlevel $indentDataStart ))
			foreach ($_user in $_snmpV3UserList)
			{
				$userName 			= $_user.userName
				$secLevel			= $_user.securityLevel
				$authProtocol 		= $_user.authProtocol
				$authPassword		= $_user.authPassword
				$privacyProtocol 	= $_user.privacyProtocol
				$privacyPassword 	= $_user.privacyPassword

				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-Prefix 	'{' 					-isItem $True				-indentlevel ($indentDataStart +1)))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-Prefix 		'snmpV3UserName' 	-value "$userName,"			-indentlevel ($indentDataStart +2)))
				if ($privacyProtocol)
				{
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-Prefix 	'v3PrivacyProtocol' -value "$privacyProtocol,"	-indentlevel ($indentDataStart +2)))
				}
				if ($authProtocol)
				{
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-Prefix 	'v3AuthProtocol' 	-value "$authProtocol," 	-indentlevel ($indentDataStart +2)))
				}	
	
				if ($authProtocol -or $privacyProtocol)
				{
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-Prefix 	'userCredentials' 								-indentlevel ($indentDataStart +2)))

					if ($authProtocol)
					{
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-Prefix 		'{' 			-isVar $True -isItem $True		-indentlevel ($indentDataStart +3)))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-Prefix 	propertyName 	-value 'v3AuthPassword,'			-indentlevel ($indentDataStart +4)))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-Prefix 	value			-value "'$authPassword'"			-indentlevel ($indentDataStart +4)))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-Prefix 		'}' 			-isItem $True					-indentlevel ($indentDataStart +3)))
					}
					if ($privacyProtocol)
					{
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-Prefix 		'{' 		-isVar $True -isItem $True		-indentlevel ($indentDataStart +3)))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-Prefix 	propertyName 	-value 'v3PrivacyPassword,'		-indentlevel ($indentDataStart +4)))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-Prefix 	value			-value "'$privacyPassword'"		-indentlevel ($indentDataStart +4)))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-Prefix 		'}' 		-isItem $True					-indentlevel ($indentDataStart +3)))	
					}
					
					
				}
				else 	# noAuth
				{
					$YMLscriptCode[-1]		= $YMLscriptCode[-1].TrimEnd() -replace ".$"
				}
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-Prefix 	'},' 					-isItem $True					-indentlevel ($indentDataStart +1)))
			}

			$YMLscriptCode[-1]		= $YMLscriptCode[-1].TrimEnd() -replace ".$"
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -Prefix 	'           ] ' 			-isItem $True						-indentlevel $indentDataStart ))

			
		}
	
		# ----- snmpTrap 
		$_snmpTrapList				= $snmpTrapList | where source -eq $ligName
		if ($_snmpTrapList)
		{
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -Prefix 	'trapDestinations: [ ' 			-isItem $True					-indentlevel $indentDataStart ))
			foreach ($_trap in $_snmpTrapList)
			{
				$format 					= $_trap.format
				$destinationAddress			= $_trap.destinationAddress
				$port 						= $_trap.port              
				$communityString 			= $_trap.communityString
				$snmpV3User 				= $_trap.userName

				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-Prefix 	'{' 					-isItem $True					-indentlevel ($indentDataStart +1)))

				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -Prefix 			'trapDestination' 	-value "$destinationAddress,"	-indentlevel ($indentDataStart +2) ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -Prefix 			'port' 				-value "$port,"					-indentlevel ($indentDataStart +2) ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -Prefix 			'trapformat' 		-value "$format,"				-indentlevel ($indentDataStart +2) ))
				if ($communityString)
				{
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -Prefix 		'communityString' 	-value "$communityString,"		-indentlevel ($indentDataStart +2) ))
				}
				if ($snmpV3User)
				{
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -Prefix 		'userName' 			-value "$snmpV3user,"			-indentlevel ($indentDataStart +2) ))
				}
				$YMLscriptCode[-1]		= $YMLscriptCode[-1].TrimEnd() -replace ".$"
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-Prefix 	'},' 					-isItem $True					-indentlevel ($indentDataStart +1)))
			}
			
			$YMLscriptCode[-1]		= $YMLscriptCode[-1].TrimEnd() -replace ".$"
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -Prefix 	'                  ]' 			-isItem $True					-indentlevel $indentDataStart ))

		}	
	
	}
	
		
	 # ---------- Generate script to file
	 YMLwriteToFile -code $YMLscriptCode -file $YMLfiles
	
}

Function Import-YMLLogicalInterconnectGroup([string]$sheetName, [string]$WorkBook, [string]$YMLfiles )
{
	$YMLscriptCode             	= [System.Collections.ArrayList]::new()
	$bayPortInfos 				= [System.Collections.ArrayList]::new()


	$cSheet, $ligSheet, $uplSheet, $snmpConfigSheet, $snmpV3UserSheet, $snmpTrapSheet 	= $sheetName.Split($SepChar)       # composer
	$ligList 					= if ($ligSheet)	 		{get-datafromSheet -sheetName $ligSheet -workbook $WorkBook				} else {$null}

	$uplList 					= if ($uplSheet)	 		{get-datafromSheet -sheetName $uplSheet -workbook $WorkBook				} else {$null}
	

	[void]$YMLscriptCode.Add((Generate-ymlheader -title 'Configure logical Interconnect Group'))

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
		$internalNetworkConsistency = if ($L.consistencyCheckingForInternalNetworks) 	{$YMLconsistencyCheckingLigEnum.Item($L.consistencyCheckingForInternalNetworks) } else {'None'}
		$interconnectConsistencyChecking = if ($L.interconnectConsistencyChecking) 		{$YMLconsistencyCheckingLigEnum.Item($L.interconnectConsistencyChecking) } else {'None'}

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

		$isNotSAS 					= $bayConfig -notlike '*SAS*'

		$_ethernetType 				= $YMLtype540Enum.item('ethernetSettings')
		


		$comment 					= '# ---------- Logical Interconnect Group {0}' 	-f $ligName
		$title 						= ' Create logical InterConnect Group {0}' 			-f $ligName	

		[void]$YMLscriptCode.Add((Generate-ymlTask 		 	-title $title -comment $comment -OVTask 'oneview_logical_interconnect_group'))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		name				-value $name							-indentlevel $indentDataStart ))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		enclosureType		-value $enclosureType					-indentlevel $indentDataStart ))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		interconnectBaySet	-value $ICBaySet						-indentlevel $indentDataStart ))	
		if ($isNotSAS)
		{
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		redundancyType		-value $redundancyType				-indentlevel $indentDataStart ))
		}

		# ------ Enclosure Index
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		enclosureIndexes											-indentlevel $indentDataStart ))
		if ($fabricModuleType -like '*FC*')
		{
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	"- -1"				-isItem $True							-indentlevel ($indentDataStart+1) ))
		}
		else 
		{
			foreach ($_index in 1..$frameCount)
			{
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix "- $_index"			-isItem $True							-indentlevel ($indentDataStart+1) ))
			}			
		}


		# ------ Ethernet Settings
		if ($isNotSAS)
		{
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		ethernetSettings														-indentlevel $indentDataStart ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			type						-value $_ethernetType 					-indentlevel ($indentDataStart +1) ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			enableIgmpSnooping			-value $enableIgmpSnooping 				-indentlevel ($indentDataStart +1) ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			enableNetworkLoopProtection	-value $enableNetworkLoopProtection 	-indentlevel ($indentDataStart +1) ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			enablePauseFloodProtection	-value $enablePauseFloodProtection		-indentlevel ($indentDataStart +1) ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			enableRichTLV				-value $enableRichTLV 					-indentlevel ($indentDataStart +1) ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			enableFastMacCacheFailover	-value $enableFastMacCacheFailover		-indentlevel ($indentDataStart +1) ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			enableStormControl			-value $enableStormControl 				-indentlevel ($indentDataStart +1) ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			stormControlPollingInterval	-value $stormControlPollingInterval 	-indentlevel ($indentDataStart +1) ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			stormControlThreshold		-value $stormControlThreshold 			-indentlevel ($indentDataStart +1) ))		
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			enableTaggedLldp			-value $enableTaggedLldp				-indentlevel ($indentDataStart +1) ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			lldpIpAddressMode			-value $lldpIpAddressMode				-indentlevel ($indentDataStart +1) ))
			if ($lldpIpv4Address)
			{
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		lldpIpv4Address				-value $lldpIpv4Address					-indentlevel ($indentDataStart +1) ))
			}
			if ($lldpIpv4Address)
			{
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		lldpIpv6Address				-value $lldpIpv6Address 				-indentlevel ($indentDataStart +1) ))
			}
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			consistencyChecking			-value $interconnectConsistencyChecking	-indentlevel ($indentDataStart +1) ))
		}


		# ------ Internal Networks
		if ($internalNetworks)
		{
			 
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix "consistencyCheckingForInternalNetworks: $internalNetworkConsistency" -isItem $True	-indentlevel $indentDataStart ))
			
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-prefix internalNetworkNames															-indentlevel $indentDataStart ))
			$internalNetArr 		= $internalNetworks.Split($SepChar)
			foreach ($_int in $internalNetArr)
			{
				$_item 				= '- {0}'	-f $_int
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		$_item					-isItem $True	-indentlevel ($indentDataStart +1) ))

			}
		}
		
		# ------ Bay Config
		if ($bayConfig)
		{
			# Get logical PortInfos in uplink sets


			$bayPortInfos =  build-portInfos -bayconfig $bayConfig 
			if ($bayPortInfos)
			{
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		"interconnectMapTemplate"						-indentlevel $indentDataStart ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			"interconnectMapEntryTemplates"				-indentlevel ($indentDataStart +1) ))
			
				foreach ($_b in $bayPortInfos)
				{
					$_frame 	= if ($_b.bayModule -like '*FC*') { -1} else {$_b.frame}

					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	   ("- permittedInterconnectTypeName: " + $_b.bayModuleName) -isItem $True	-indentlevel ($indentDataStart + 2)  ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		"  enclosureIndex"					-value $_frame						-indentlevel ($indentDataStart + 2)  ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		"  logicalLocation"														-indentlevel ($indentDataStart + 2)  ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			"  locationEntries"													-indentlevel ($indentDataStart + 3)  ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 				"- type"					-value Bay							-indentlevel ($indentDataStart + 4)  ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 				"  relativeValue"			-value $_b.bay						-indentlevel ($indentDataStart + 4)  ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 				"- type"					-value Enclosure					-indentlevel ($indentDataStart + 4)  ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 				"  relativeValue"			-value $_frame						-indentlevel ($indentDataStart + 4)  ))

				}		
			}
		}

	}
	
	 # ---------- Generate script to file
	 YMLwriteToFile -code $YMLscriptCode -file $YMLFiles
}

Function Import-LogicalSwitchGroup([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $PSscriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
		
	foreach ( $swg in $List)
	{
	
		$name                   = $swg.name
		$switchType 			= $swg.switchType
		$numberofSwitches 		= $swg.numberofSwitches

		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating Logical Switch Group {0} "' -f $name) -isVar $False ))
 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'swg' 						-Value ("get-OVLogicalSwitchGroup -name '{0}' -ErrorAction SilentlyContinue " -f $name ) ))		#HKD

		ifBlock			-condition 'if ($swg -eq $Null)' 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'switchType' 					-Value ('Get-OVSwitchType -name "{0}"'	-f $switchType) -indentlevel 1))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('new-OVLogicalSwitchGroup -name "{0}" -switchType $switchType -NumberOfSwitches {1}' -f $name,$numberofSwitches) -isVar $False -indentlevel 1))
	
		endBlock

		# Skip creating because resource already exists
		elseBlock
		[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW ' + "{0} does not exist or " -f $name   ) -isVar $False -indentlevel 1 ))
		endElseBlock

		newLine

		# --- Scopes
		if ($scopes)
		{
			[void]$PSscriptCode.Add( (Generate-PSCustomVarCode -Prefix 'object' -Value ('Get-OVLogicalSwitchGroup -name "{0}" -ErrorAction SilentlyContinue' -f $name) -indentlevel 1)) #HKD
			generate-scopeCode -scopes $scopes -indentlevel 1
			newLine

		}
	}
	
	if ($List)
	{
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	}
	
	 # ---------- Generate script to file
	 writeToFile -code $PSscriptCode -file $ps1Files


}


Function Import-LogicalSwitch([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $PSscriptCode             	= [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	
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

			

		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating Logical Switch {0} "' -f $name) -isVar $False ))
 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'lsw' 						-Value ("get-OVLogicalSwitch -name '{0}' -ErrorAction SilentlyContinue " -f $name ) ))		#HKD

		ifBlock			-condition 'if ($lsw -eq $Null)' 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'logicalSwitchGroup' 			-Value ('Get-OVLogicalSwitchGroup -name "{0}" -ErrorAction SilentlyContinue'	-f $logicalSwitchGroup) -indentlevel 1))
		
		$namegroupParam = ' -name "{0}" -logicalSwitchGroup $logicalSwitchGroup ' -f $name
		$managedParam 	= if ($isManaged) { ' -Managed'} else {' -Monitored'}

		$s1 			= if ($switch1Address) { ' -Switch1Address {0} ' -f $switch1Address} else {''}
		$s2 			= if ($switch2Address) { ' -Switch2Address {0} ' -f $switch2Address} else {''}
		$addressParam 	= $s1 + $s2

		generate-credentialCode -password $sshPassword -username $sshUserName -component 'LOGICAL SWITCH'-indentLevel 1 -PSscriptCode $PSscriptCode
		$credParam 		= ' -sshUserName "{0}" -sshPassword $securePassword' -f $sshUserName # $securePassword is defined in  generate-credentialCode

		 
		if ($issnmpV3)
		{
			$snmpParam 	= ' -snmpV3 $True -SnmpUserName "{0}" -SnmpAuthLevel "{1}" ' -f $snmpV3User, $snmpAuthLevel
			if ($snmpAuthLevel -eq 'Auth')
			{	
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'authPassword' -value ('"{0}" | ConvertTo-SecureString -AsPlainText -Force' -f $snmpAuthPassword) -indentLevel 1 ))
				$snmpParam	+= ' -snmpAuthProtocol "{0}" -snmpAuthPassword $authPassword ' -f  $snmpAuthProtocol
			}
			else
			{
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'authPassword' -value ('"{0}" | ConvertTo-SecureString -AsPlainText -Force' -f $snmpAuthPassword) -indentLevel 1 ))
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'privPassword' -value ('"{0}" | ConvertTo-SecureString -AsPlainText -Force' -f $snmpPrivPassword)  -indentLevel 1 ))
				$snmpParam	+=  ' -snmpAuthProtocol "{1}" -snmpAuthPassword $authPassword -snmpPrivProtocol "{2}" -snmpPrivPassword $privPassword' -f $snmpAuthProtocol, $snmpPrivProtocol			
			}
		}
		else # snmpV1
		{
			$snmpParam 	= ' -snmpV1 $True -snmpPort {0} -snmpCommunity "{1}" ' -f $snmpPort, $snmpCommunity 
		}
		
		
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('new-OVLogicalSwitch {0}{1} `' -f $namegroupParam, $managedParam) -isVar $False  -indentlevel 1))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0}{1} `' 						-f $addressParam, $credParam) 	   -isVar $False  -indentlevel 4))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0}' 							-f $snmpParam) 					   -isVar $False  -indentlevel 4))
		newLine
	
		endBlock

		# Skip creating because resource already exists
		elseBlock
		[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW ' + "{0} does not exist or " -f $name   ) -isVar $False -indentlevel 1 ))
		endElseBlock

		newLine

	}
	
	if ($List)
	{
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	}
	
	 # ---------- Generate script to file
	 writeToFile -code $PSscriptCode -file $ps1Files


}


Function Import-SANmanager([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $PSscriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	
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


			

		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating SAN Manager {0} "' -f $name) -isVar $False ))
 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'san' 						-Value ("get-OVSANManager -name '{0}' -ErrorAction SilentlyContinue " -f $name ) ))	#HKD	
		
		ifBlock			-condition 'if ($san -eq $Null)' 		
		$nameParam 		= ' -HostName "{0}" -type "{1}" ' -f $name, $type
		if ( ($userName) -and ($password) )
		{
			generate-credentialCode -password $password -username $userName -component 'SAN MANAGER'-indentLevel 1 -PSscriptCode $PSscriptCode
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
						[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'authPassword' -value ('"{0}" | ConvertTo-SecureString -AsPlainText -Force' -f $snmpAuthPassword) -indentLevel 1 ))
						$authParam	+= ' -snmpAuthProtocol "{0}" -snmpAuthPassword $authPassword ' -f  $snmpAuthProtocol
					}
				'AuthAndPriv'
					{
						[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'authPassword' -value ('"{0}" | ConvertTo-SecureString -AsPlainText -Force' -f $snmpAuthPassword) -indentLevel 1 ))
						[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'privPassword' -value ('"{0}" | ConvertTo-SecureString -AsPlainText -Force' -f $snmpPrivPassword)  -indentLevel 1 ))
						$authParam	+=  ' -snmpAuthProtocol "{1}" -snmpAuthPassword $authPassword -snmpPrivProtocol "{2}" -snmpPrivPassword $privPassword' -f $snmpAuthProtocol, $snmpPrivProtocol			
					}
			}	
		}

		
		
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('new-OVSANManager {0} `' 	-f $nameParam) -isVar $False  -indentlevel 1))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0}' 						-f $authParam) -isVar $False  -indentlevel 3))
		newLine
	
		endBlock

		# Skip creating because resource already exists
		elseBlock
		[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW ' + "{0} does not exist." -f $name   ) -isVar $False -indentlevel 1 ))
		endElseBlock

		newLine

	}
	
	if ($List)
	{
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	}
	
	 # ---------- Generate script to file
	 writeToFile -code $PSscriptCode -file $ps1Files


}



Function Import-StorageSystem([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $PSscriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	
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



			

		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating Storage System {0} "' -f $name) -isVar $False ))
 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'sts' 						-Value ("get-OVStorageSystem -name '{0}' -ErrorAction SilentlyContinue " -f $name ) )) #HKD		
		ifBlock 'if ($sts -eq $Null)' 	
		
		$s 					= if ($showSystemDetails) { ' -ShowSystemDetails'} else {''}
		$nameParam 			= (' -HostName "{0}" -Family "{1}" ' -f $name, $family) + $s


		$authParam 			= $null
		if ( ($userName) -and ($password) )
		{
			generate-credentialCode -password $password -username $userName -component 'STORAGE SYSTEM'-indentLevel 1 -PSscriptCode $PSscriptCode
			$authParam 		= ' -Credential $cred'  		# $cred is defined in  generate-credentialCode
		}

		if ($vips)
		{
			$ip , $netName 	= $vips.Split($Equal)
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'net' 	-Value ("get-OVNetwork -name '{0}' -ErrorAction SilentlyContinue " -f $netName.trim() ) -indentlevel 1 ))	 #HKD
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'vips' 	-Value ('@{' + ('"{0}" = $net' -f $ip.trim())  + '}') 							-indentlevel 1 )) 
			$vipsParam 			 = ' -VIPS $vips'
		}
		
		
		$domainParam 		= if ($domain) { ' -Domain "{0}" ' -f $domain } else {''}
		
		
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('new-OVStorageSystem {0}{1} `' 	-f $nameParam, $authParam) 	-isVar $False  -indentlevel 1))

		if ($family -eq 'StoreServ')
		{
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0} | wait-OVTaskComplete' 	-f $domainParam) 			-isVar $False  -indentlevel 3))

		}
		else
		{
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0} | wait-OVTaskComplete'	-f $vipsParam) 				-isVar $False  -indentlevel 3))
		}
		newLine

		if ($storagePool)
		{
			$pool 		= "@('" + $storagePool.replace($sepChar, "'$comma'") + "')"
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'pool' 						-Value ("{0}" -f $pool ) -indentlevel 1 ))	
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'sts' 						-Value ("get-OVStorageSystem | where name -eq  '{0}' " -f $name ) -indentlevel 1 ))	
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'new-OVStoragePool -pool $pool -StorageSystem $sts | wait-OVTaskComplete'  -isVar $False  -indentlevel 1))
			# Set pool to managed state
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '$pool | % {get-OVStoragePool -name $_ | set-OVStoragePool -Managed $true }'  -isVar $False  -indentlevel 1))
			
			newLine
		}

		endIfBlock
		# Skip creating because resource already exists
		elseBlock
		[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW ' + "{0} does not exist." -f $name   ) -isVar $False -indentlevel 1 ))
		endElseBlock

		newLine

	}
	
	if ($List)
	{
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	}
	
	 # ---------- Generate script to file
	 writeToFile -code $PSscriptCode -file $ps1Files


}


Function set-Param([string]$var, [string]$value)
{
	$param 				= if ($value -eq 'True') { " -$var "} else {""}
	return $param

}
Function Import-StorageVolumeTemplate([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $PSscriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	
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
			

		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating Storage Volume Template {0} "' -f $name) -isVar $False ))
 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'svt' 						-Value ("get-OVStorageVolumeTemplate -name '{0}' -ErrorAction SilentlyContinue" -f $name ) ))		#HKD
		
		ifBlock			-condition 'if ($svt -eq $Null)' 		
		$descParam 			= if ($description) { ' -Description "{0}" ' -f $description} else {''}
		$nameParam 			= (' -Name "{0}" {1} ' -f $name, $descParam) 

		if ($storagePool)
		{	
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'stp' 						-Value ("get-OVStoragePool -name  '{0}' -ErrorAction SilentlyContinue " -f $storagePool ) -indentlevel 1 ))	#HKD
			
			ifBlock		-condition 'if ($stp -ne $Null)' 	-indentlevel 1
			# ---- Storage Pool and SnapshotStoragePool
			$storagePoolParam 			= ' -StoragePool $stp ' + $lockStoragePool
			if ($snapshotStoragePool)
			{
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'sstp' 					-Value ("get-OVStoragePool -name '{0}' -ErrorAction SilentlyContinue " -f $snapshotStoragePool ) -indentlevel 2 ))	#HKD
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'sstp' -value 'if ($sstp -ne $Null) { $sstp } else {$stp}'  -indentlevel 2) )
				$storagePoolParam 		+= ' -SnapshotStoragePool $sstp ' + $lockSnapshotStoragePool
			}

			# ---- Storage System
			if ($storageSystem)		
			{
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ' # ----- Get Storage System' -isVar $False -indentlevel 2 ))
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'sts' 						-Value ("get-OVStorageSystem -name '{0}' -ErrorAction SilentlyContinue " -f $storageSystem ) -indentlevel 2 ))	#HKD
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
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'folder' 	-Value ("(get-OVStoragePool -name default).deviceSpecificAttributes.Folders | where name -eq  '{0}' " -f $folder ) -indentlevel 2 ))	
				$a12 					= ' -Folder $folder' + $lockFolder													
			}

			$a13						= $null 
			if ($performancePolicy)
			{
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'performancePolicy' 	-Value ('Show-OVStorageSystemPerformancePolicy -InputObject $sts -name "{0}" ' -f $performancePolicy ) -indentlevel 2 ))	
				$a13 					= ' -PerformancePolicy $performancePolicy' + $lockPerformancePolicy	
			}
			
			$a14						= $null 
			if ($volumeSet)
			{
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'volumeSet' 	-Value ("get-OVStorageVolumeSet -name '{0}' -ErrorAction SilentlyContinue " -f $volumeSet ) -indentlevel 2 ))	#HKD
				$a14 					= ' -VolumeSet $volumeSet' + $lockVolumeSet
			}
			$attributes4Param			= $a12 + $a13 + $a14

			#---- code here
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('new-OVStorageVolumeTemplate {0}{1}{2} `'	-f $nameParam, $capacityParam, $storagePoolParam )	-isVar $False  -indentlevel 2))
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0}{1} `'									-f $attributes1Param, $attributes2Param) 			-isVar $False  -indentlevel 3))
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0} `'										-f $attributes3Param) 								-isVar $False  -indentlevel 3))
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0}'										-f $attributes4Param) 								-isVar $False  -indentlevel 3))
			
			newLine

			# --- Scopes
			if ($scopes)
			{
				newLine
				[void]$PSscriptCode.Add( (Generate-PSCustomVarCode -Prefix 'object' -Value ('getHPOVStorageVolumeTemplate -name "{0}" -ErrorAction SilentlyContinue ' -f $name) -indentlevel 1))   #HKD
				generate-scopeCode -scopes $scopes -indentlevel 1

			}

			endifBlock 		-condition 'if ($stp -ne $Null)'  -indentlevel 1

			elseBlock 	-indentlevel 1
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW "storage pool {0} does not exist. Cannot create volume template"' -f $storagepool) 		-isVar $False -indentlevel 2) )
			newLine
			endElseBlock -indentlevel 1


			endIfBlock  -condition 'if $svt -eq $Null'
			elseBlock
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW "volume template {0} already exists. skip creating volume template" ' -f $name) 	-isVar $False -indentlevel 1) )	
			endElseBlock



		}
		else {}#TBD 



	}
	
	if ($List)
	{
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	}
	
	 # ---------- Generate script to file
	 writeToFile -code $PSscriptCode -file $ps1Files


}


Function Import-StorageVolume([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $PSscriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	
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
			

		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating Storage Volume {0} "' -f $name) -isVar $False ))
 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'vol' 						-Value ("get-OVStorageVolume -name '{0}' -ErrorAction SilentlyContinue " -f $name ) ))		#HKD
		
		ifBlock 		-condition  'if ($Null -eq $vol)' 		
		$descParam 				= if ($description) { ' -Description "{0}" ' -f $description} else {''}
		$nameParam 				= (' -Name "{0}" {1} ' -f $name, $descParam) 

		if ($volumeTemplate)
		{
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'svt' 					-Value ("get-OVStorageVolumeTemplate -name '{0}' -ErrorAction SilentlyContinue " -f $volumeTemplate ) -indentlevel 1 )) #HKD
			ifBlock -condition  'if ($Null -ne $svt)' -indentlevel 1
			$volumeTemplateParam 	= ' -VolumeTemplate $svt'

			#---- code here
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('new-OVStorageVolume {0}{1}'	-f $nameParam, $volumeTemplateParam )	-isVar $False  -indentlevel 2))
			
			newLine

			# --- Scopes
			if ($scopes)
			{
				newLine
				[void]$PSscriptCode.Add( (Generate-PSCustomVarCode -Prefix 'object' -Value ('get-OVStorageVolume -name "{0}" -ErrorAction SilentlyContinue ' -f $name) -indentlevel 1)) #HKD
				generate-scopeCode -scopes $scopes -indentlevel 1

			}
			endIfBlock  -indentlevel 1
			elseBlock   -indentlevel 1
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW "volume template {0} not found. skip creating volume" ' -f $volumeTemplate) 	-isVar $False -indentlevel 2) )
			endElseBlock -indentlevel 1
			
		}
		else # standalone volume no template
		{
			if ($storagePool)
			{	
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'stp' 						-Value ("get-OVStoragePool -name '{0}' -ErrorAction SilentlyContinue " -f $storagePool ) -indentlevel 1 ))	#HKD
				ifBlock -condition 'if ( $Null -ne $stp)' 		-isVar $False -indentlevel 1

				# ---- Storage Pool and SnapshotStoragePool
				$storagePoolParam 			= ' -StoragePool $stp ' + $lockStoragePool
				if ($snapshotStoragePool)
				{
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'sstp' 					-Value ("get-OVStoragePool -name '{0}' -ErrorAction SilentlyContinue " -f $snapshotStoragePool ) -indentlevel 2 ))	#HKD
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'sstp' -value 'if ($sstp -ne $Null) { $sstp } else {$stp}'  -indentlevel 2) )
					$storagePoolParam 		+= ' -SnapshotStoragePool $sstp ' + $lockSnapshotStoragePool
				}

				# ---- Storage System
				if ($storageSystem)		
				{
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ' # ----- Get Storage System' -isVar $False -indentlevel 2 ))
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'sts' 						-Value ("get-OVStorageSystem -name '{0}' -ErrorAction SilentlyContinue " -f $storageSystem ) -indentlevel 2 ))	#HKD
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
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'folder' 	-Value ("(get-OVStoragePool -name default).deviceSpecificAttributes.Folders | where name -eq  '{0}' " -f $folder ) -indentlevel 2 ))	
					$a12 					= ' -Folder $folder' + $lockFolder													
				}

				$a13						= $null 
				if ($performancePolicy)
				{
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'performancePolicy' 	-Value ('Show-OVStorageSystemPerformancePolicy -InputObject $sts -name "{0}" ' -f $performancePolicy ) -indentlevel 2 ))	
					$a13 					= ' -PerformancePolicy $performancePolicy' + $lockPerformancePolicy	
				}
				
				$a14						= $null 
				if ($volumeSet)
				{
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'volumeSet' 	-Value ("get-OVStorageVolumeSet -name '{0}' -ErrorAction SilentlyContinue " -f $volumeSet ) -indentlevel 2 ))	#HKD
					$a14 					= ' -VolumeSet $volumeSet' + $lockVolumeSet
				}
				$attributes4Param			= $a12 + $a13 + $a14

				#---- code here
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('new-OVStorageVolume {0}{1}{2} `'	-f $nameParam, $capacityParam, $storagePoolParam )	-isVar $False  -indentlevel 2))
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0}{1} `'									-f $attributes1Param, $attributes2Param) 			-isVar $False  -indentlevel 3))
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0} `'										-f $attributes3Param) 								-isVar $False  -indentlevel 3))
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0}'										-f $attributes4Param) 								-isVar $False  -indentlevel 3))
				
				newLine

				# --- Scopes
				if ($scopes)
				{
					newLine
					[void]$PSscriptCode.Add( (Generate-PSCustomVarCode -Prefix 'object' -Value ('get-OVStorageVolume -name "{0}" -ErrorAction SilentlyContinue ' -f $name) -indentlevel 1)) #HKD
					generate-scopeCode -scopes $scopes -indentlevel 1

				}

				endBlock 	-indentlevel 1

				elseBlock 	-indentlevel 1
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW "storage pool {0} does not exist. Cannot create volume template"' -f $storagepool) 		-isVar $False -indentlevel 2) )
				newLine
				endElseBlock -indentlevel 1

			}
			else {}#TBD 
		}

		endIfBlock # end check on if $vol -eq $Null

		elseBlock
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW "volume {0} already exists. skip creating volume" ' -f $name) 	-isVar $False -indentlevel 1) )
		endElseBlock




	}
	
	if ($List)
	{
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	}
	
	 # ---------- Generate script to file
	 writeToFile -code $PSscriptCode -file $ps1Files


}


Function Import-LogicalJBOD([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $PSscriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	
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



		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '# ----------------------------------------------------------------'  -isVar $False ))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating Logical JBOD {0} "' -f $name) -isVar $False ))
 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'jbod' 						-Value ("get-OVLogicalJBOD -name '{0}' -ErrorAction SilentlyContinue " -f $name ) ))		#HKD
		
		ifBlock 		-condition	'if ($Null -eq $jbod )' 		
		$descParam 				= if ($description) { ' -Description "{0}" ' -f $description} else {''}
		$nameParam 				= (' -Name "{0}" {1} ' -f $name, $descParam) 
		$eraseParam 			= if ($eraseDataOnDelete -eq 'True') { ' -eraseDataOnDelete $True'} else {''}

		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'driveEnclosure' 				-Value ("Get-OVDriveEnclosure -name '{0}' -ErrorAction SilentlyContinue " -f $driveEnclosure) -indentlevel 1 )) #HKD
		
		ifBlock 		-condition 'if ( $Null -ne $driveEnclosure )'  -indentlevel 1
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('New-OVLogicalJbod	{0} `' 									-f $nameParam) 										-isVar $False -indentlevel 2)) 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix  ' -InputObject $driveEnclosure `'  																				-isVar $False -indentlevel 7)) 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix (' -NumberofDrives {0} -MinDriveSize {1} -MaxDriveSize {2} `'	-f $numberofDrives, $minDriveSize , $MaxDriveSize)	-isVar $False -indentlevel 7)) 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix (' -DriveType {0} {1} '										-f $driveType, $eraseParam ) 						-isVar $False -indentlevel 7)) 
			
		endIfBlock   	-indentlevel 1 # end of check drive enclosure null

		elseBlock 		-indentlevel 1
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW "No such drive enclosure {0} to create JBOD."' -f $driveEnclosure)	-isVar $False -indentlevel 2) )
		endElseBlock 	-indentlevel 1

		endIfBlock 		-condition 'if $Null -eq $jbod '
		elseBlock
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW "logical JBOD {0} already exists. skip creating JBOD" ' -f $name) 	-isVar $False -indentlevel 1) )
		endElseBlock



	}



	
	if ($List)
	{
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	}

	# ---------- Generate script to file
	writeToFile -code $PSscriptCode -file $ps1Files


}


# ---------- Enclosure Group
Function Import-EnclosureGroup([string]$sheetName, [string]$WorkBook, [string]$ps1Files )
{
    $PSscriptCode             = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer

	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
	
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

		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating enclosure Group {0} "' -f $name) -isVar $False ))
 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'eg' 						-Value ("get-OVEnclosureGroup -name '{0}' -ErrorAction SilentlyContinue " -f $name ) ))		#HKD

		ifBlock			-condition 'if ($Null -eq $eg)' 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('# -------------- Attributes for enclosure group {0} ' -f $name) -isVar $False -indentlevel 1))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'name' 						-Value ('"{0}"'	-f $name) -indentlevel 1))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'enclosureCount' 				-Value ('{0}'	-f $enclosureCount) -indentlevel 1))

		# --- IP v4 addressing Mode
		$v4AddressPoolParam = $null
		if ([string]::IsNullOrWhiteSpace($IPv4AddressingMode) )
		{
			$ipV4AddressingMode = 'DHCP'
		}
		$ipV4AddressingMode			= $ipV4AddressingMode.Trim()
		$v4AddressPoolParam 		=   ' -IPv4AddressType $ipV4AddressType'
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'ipV4AddressType' 			-Value ('"{0}"'	-f $ipV4AddressingMode) -indentlevel 1))

		
		if ($ipV4AddressingMode -eq 'AddressPool')
		{
			$v4Range 				= 	"@('" + $ipV4Range.replace($SepChar, "','") + "')"
			$value 					= $v4Range + ' | % {Get-OVAddressPoolRange -name  $_ -ErrorAction SilentlyContinue } ' #HKD
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'IPv4AddressRange' 		-Value $value -indentlevel 1))

			$v4AddressPoolParam 	+=   ' -IPv4AddressRange $IPv4AddressRange'
		}

		# --- IP v6 addressing Mode
		$v6AddressPoolParam = $null
		if ($ipV6AddressingMode)
		{
			$ipV6AddressingMode			= $ipV6AddressingMode.Trim()
			$v6AddressPoolParam 		=   ' -IPv6AddressType $ipV6AddressType'
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'ipV6AddressType' 			-Value ('"{0}"'	-f $ipV6AddressingMode) -indentlevel 1))

			
			if ($ipV6AddressingMode -eq 'AddressPool')
			{
				$v6Range 				= 	"@('" + $ipV6Range.replace($SepChar, "','") + "')"
				$value 					= $v6Range + ' | % {Get-OVAddressPoolRange -name $_  -ErrorAction SilentlyContinue} ' #HKD
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'IPv6AddressRange' 		-Value $value -indentlevel 1))
	
				$v6AddressPoolParam 	+=   ' -IPv6AddressRange $IPv6AddressRange'
			}
		}

		# Power Mode
		if ($powerMode)
		{
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'powerMode' -Value ('"{0}"' -f $powerMode) -indentlevel 1))
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
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix $variable -Value $value -indentlevel 1))

				$ligGroupMapping 	= $ligGroupMapping.replace($varName, $variable)
				$i++
			}
			 
			# 2- We build the hash table
			
			$ligGroupMapping	= '@{' + $ligGroupMapping.replace($sepChar,';') + '}'		# Hash Table
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'ligGroupMapping' -Value $ligGroupMapping -indentlevel 1))

			$ligMappingParam 	= ' -LogicalInterconnectGroupMapping $ligGroupMapping'

		}

		# OSdeployment
		$deploymentTypeParam 	= ''
		switch ($deploymentMode)
		{
			'External' 
				{
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'deployNetwork' -Value ('Get-OVNetwork -name "{0}"' -f $deploymentNetwork) -indentlevel 1))
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
		[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix $( 'New-OVEnclosureGroup	-Name $name -enclosureCount $enclosureCount {0}{1} `' -f $v4AddressPoolParam, $v6AddressPoolParam) -isVar $False -indentlevel 1))
		[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('{0}{1} `' 	-f $ligMappingParam, $powerModeParam) -isVar $False -indentlevel 6))
		if ($deploymentTypeParam)
		{
			[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('{0} `'	-f $deploymentTypeParam) -isVar $False -indentlevel 6))
		}
		newLine # to end the command
		
		# --- Scopes
		if ($scopes)
		{
			newLine
			[void]$PSscriptCode.Add( (Generate-PSCustomVarCode -Prefix 'object' -Value 'get-OVEnclosureGroup -name $name -ErrorAction SilentlyContinue ' -indentlevel 1))   #HKD
			generate-scopeCode -scopes $scopes -indentlevel 1

		}


		endIfBlock

		# Skip creating because resource already exists
		elseBlock
		[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW {0} already exists.' -f $name ) -isVar $False -indentlevel 1 ))
		endElseBlock

		newLine
        
        
	}
	
	if ($List)
	{
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	}
	
	 # ---------- Generate script to file
	 writeToFile -code $PSscriptCode -file $ps1Files

}

Function Import-YMLEnclosureGroup([string]$sheetName, [string]$WorkBook, [string]$YMLFiles )
{
    $YMLscriptCode          = [System.Collections.ArrayList]::new()
	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	
	[void]$YMLscriptCode.Add((Generate-ymlheader -title 'Configure Enclosure Group'))
	[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix set_fact 		-isVar $True 										-indentlevel 1 ))
	[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	IC_OFFSET	-value 3											-indentlevel 2 ))
			

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
		
		#---- Interconnect Bay mapping
		if ($ligMapping)
		{
			$comment 					= '# ---------- Enclosure Group {0}' 			-f $name
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	$comment	-isItem $True						-indentlevel 1 ))

			# 1 - Build array of unique LIG
			$vars 				= $ligMapping  -replace 'Frame\d+\s+=\s+' , ''			# Remove Framex=
			$vars 				= $vars.replace($SepChar, $Comma)
			$varArray			= $vars.Split($Comma)								# Build array of variable names
			$varArray 			= $varArray | sort -Unique							# Get unique value


			foreach ($_ligName in $varArray)
			{
				$title 						= ' Get lig {0} Information' 			-f $_ligName	
				[void]$YMLscriptCode.Add((Generate-ymlTask 	-title $title -isData $False 	-OVTask 'oneview_logical_interconnect_group_facts'))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix name			-value $_ligName									-indentlevel 2 ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix set_fact 		-isVar $True 										-indentlevel 1 ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	lig			-value "'{{logical_interconnect_groups}}' "			-indentlevel 2 ))
				
				$title 						= ' Try SAS lig {0} Information' 			-f $_ligName	
				[void]$YMLscriptCode.Add((Generate-ymlTask 	-title $title -isData $False 	-OVTask 'oneview_sas_logical_interconnect_group_facts'))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix name			-value $_ligName									-indentlevel 2 ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix set_fact 		-isVar $True 										-indentlevel 1 ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	lig			-value "'{{sas_logical_interconnect_groups}}' "		-indentlevel 2 ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix when 			-value "(lig|length == 0)" 							-indentlevel 1 ))

				$_name 			= $_ligName.trim().replace($Space, '_').replace('-','_')
				$var_uri 		= "var_{0}_uri"				-f $_name
				$var_primary 	= "var_{0}_bay_primary"		-f $_name
				$var_secondary 	= "var_{0}_bay_secondary"	-f $_name

				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix set_fact 			-isVar $True 											-indentlevel 1 ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	$var_uri		-value "'{{lig[0].uri}}' "								-indentlevel 2 ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	$var_primary	-value "'{{lig[0].interconnectBaySet}}' "				-indentlevel 2 ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	$var_secondary	-value "'{{lig[0].interconnectBaySet + IC_OFFSET}}' "	-indentlevel 2 ))
				
			}
		}

		#Subnet mapping
		if ($ipV4Range)
		{
			$rangeArr 			= $ipV4Range.Split($SepChar)
			foreach ($_range in $rangeArr)
			{
				$_range			= $_range.Trim() 	
				$var_subnet_uri = "var_{0}_subnet_uri" 	-f $_range
				$title			= ' Get subnet URI for subnet {0}' 			-f $_range	
				[void]$YMLscriptCode.Add((Generate-ymlTask 	-title $title -isData $False 		-OVTask 'oneview_id_pools_ipv4_range_facts'))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix name				-value $_range								-indentlevel 2 ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix set_fact 			-isVar $True 								-indentlevel 1 ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	$var_subnet_uri	-value "'{{id_pools_ipv4_ranges[0].uri}}'"	-indentlevel 2 ))
			}
		}

			

		# Create enclosure group
		

		$title 						= ' Create Enclosure group' 			-f $name	
		[void]$YMLscriptCode.Add((Generate-ymlTask 	-title $title 	-OVTask 'oneview_enclosure_group'))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	name						-value $name					-indentlevel $indentDataStart ))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	enclosureCount				-value $enclosureCount			-indentlevel $indentDataStart ))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	powerMode					-value $powerMode				-indentlevel $indentDataStart ))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	ipv6AddressingMode			-value $ipV6AddressingMode		-indentlevel $indentDataStart ))
		if ($ligMappingArr)
		{
			$ligMappingArr 				= $ligMapping.split($sepChar)
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	interconnectBayMappings										-indentlevel $indentDataStart ))
			foreach ($l in $ligMappingArr)
			{
				$frame,$lig_list	= $l.split($Equal)	# Frame1=LIG-ETH,LIG-SAS,LIG-FC
				$frame 				= $frame.Trim().ToLower()
				$_frame 			= $frame.Trim().replace('frame','')
				$ligArr 			= $lig_list.Split($COMMA)
				foreach ($_ligName in $ligArr)
				{
					$_name 			= $_ligName.trim().replace($Space, '_').replace('-','_')
					$var_uri 		= "var_{0}_uri"				-f  $_name
					$var_primary 	= "var_{0}_bay_primary"		-f  $_name
					$var_secondary 	= "var_{0}_bay_secondary"	-f  $_name
					newline -code $YMLscriptCode
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	interconnectBay				-value "'{{$var_primary}}'" -isVar $True	-indentlevel ($indentDataStart+1) ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	logicalInterconnectGroupUri	-value "'{{$var_uri}}'"						-indentlevel ($indentDataStart+1) ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	enclosureIndex				-value $_frame 								-indentlevel ($indentDataStart+1) ))
					
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	interconnectBay				-value "'{{$var_secondary}}'" -isVar $True	-indentlevel ($indentDataStart+1) ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	logicalInterconnectGroupUri	-value "'{{$var_uri}}'"						-indentlevel ($indentDataStart+1) ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	enclosureIndex				-value $_frame								-indentlevel ($indentDataStart+1) ))
			
				}

			}
		}


		# --- Subnet
		$ipV4AddressingMode 		= $ipV4AddressingMode.replace('AddressPool', 'IpPool')
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	ipAddressingMode			-value $ipV4AddressingMode		-indentlevel $indentDataStart ))
		
		if ($ipV4Range)
		{
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix ipRangeUris													-indentlevel $indentDataStart ))
			$rangeArr 			= $ipV4Range.Split($SepChar)
			foreach ($_range in $rangeArr)
			{
				$_range			= $_range.Trim() 	
				$var_subnet_uri = "var_{0}_subnet_uri" 	-f $_range
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix "'{{$var_subnet_uri}}'"  -isVar $True 	-isItem $True	-indentlevel ($indentDataStart +2) ))

			}

		}

		newLine -code $YMLscriptCode
		

	}

	 # ---------- Generate script to file
	 YMLwriteToFile -code $YMLscriptCode -file $YMLfiles

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
		$enclosureNewName 	= $le.enclosureNewName
		$enclosureGroup		= $le.enclosureGroup
		$manualAddresses 	= $le.manualAddresses
		$fwBaseline 		= $le.firmwareBaseline
		$fwInstall     		= $le.forceInstallFirmware
		$scopes			    = $le.scopes

		# Create logicalEnclosure filename here per LE
		$filename 			= "$subdir\" + $name.Trim().Replace($Space, '') + '.ps1'
		[void]$ListPS1files.Add($filename)

		$PSscriptCode         = [System.Collections.ArrayList]::new()
		connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode

		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating logical enclosure {0} "' -f $name) -isVar $False ))
 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'le' 						-Value ("get-OVLogicalEnclosure -name '{0}' -ErrorAction SilentlyContinue " -f $name ) ))		#HKD
		
		ifBlock			-condition 'if ($le -eq $Null)' 	
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('# -------------- Attributes for logical enclosure {0} ' -f $name) -isVar $False -indentlevel 1))
		

		# rename enclosure
		if ( ($enclosureNewName) -and ($enclosureName -ne $enclosureNewName))		# If there are neww names
		{
			$SNArray 			= "@('" + $enclosureSN.Replace($SepChar,"'$Comma'") 	+ "')"
			#$nameArray 			= "@('" + $enclosureName.Replace($SepChar,"'$Comma'")   + "')"
			$newNameArray 		= "@('" + $enclosureNewName.Replace($SepChar,"'$Comma'")   + "')"

			newLine
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '# -- Renaming enclosures  ' -isVar $False -indentlevel 1))
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix serialNumbers  	-value $SNArray -indentlevel 1))
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix enclosureNewNames  -value $newNameArray -indentlevel 1))

			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'for ($i=0; $i -lt $serialNumbers.Count; $i++)' -isVar $False -indentlevel 1))
			startBlock -indentlevel 1
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'this_enclosure' -Value 'Get-OVEnclosure | where serialNumber -Match $serialNumbers[$i]'  -indentlevel 2))
			ifBlock -condition 'if ($this_enclosure)   ' 		-isVar $False -indentlevel 2
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'Set-OVEnclosure -inputObject $this_enclosure -Name $enclosureNewNames[$i]' -isVar $False -indentlevel 3))	
			endIfBlock -indentlevel 2
			endBlock -indentlevel 1
		}

		# --- Enclosure		
		$enclosure 				= $enclosureSN.Split($SepChar)[0]
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix enclosure -value ("Get-OVEnclosure | where serialNumber -match '{0}' " -f $enclosure) -indentlevel 1 ))
		$enclParam 				= ' -Enclosure $enclosure'

		# --- EnclosureGroup
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix enclosureGroup -value ("Get-OVEnclosureGroup -name '{0}' " -f $enclosureGroup) -indentlevel 1 ))
		$egParam 				= ' -EnclosureGroup $enclosureGroup'


		# fwBaseline
		$fwParam 				= $Null
		if ($fwBaseline)
		{
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix fwBaseline -value ("Get-OVBaseline -SPPname '{0}' " -f $fwBaseline) -indentlevel 1 ))
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix fwInstall -value ('${0}' -f $fwInstall) -indentlevel 1 ))
			$fwParam = ' -FirmwareBaseline $fwBaseline -ForceFirmwareBaseline $fwInstall'
		}

		#HKD
		# manualAdddresses
		if ($manualAddresses)
		{
			$ebipaParam = $Null
			$frameArr			= $manualAddresses.Split($Delimiter)		# '\' 
			$frameArr			= $frameArr -replace $CR, ''				# Remnove extra line

			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'ebipa' 				-Value '@{' 					-indentlevel 1))

			foreach ($_fr in $frameArr)
			{
				$_frame, $_bay 	= $_fr.Split('@')
				$_frame 		= $_frame -replace '=', ''
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix "$_frame = @{" 				-isVar $False		-indentlevel 2))
				# Transform the device and Interconnect string
				$_f 			= $_bay.IndexOf('{') + 1
				$_l 			= $_bay.LastIndexof('}')
				$_len 			= $_l - $_f
				$_bay 			= $_bay.Substring($_f, $_len)  # Remove { }
				$_bayArr		= $_bay.Split(';')
				foreach ($_b in  $_bayArr)
				{
					$_b 		= $_b.Replace('={', '=@{')    # Device8 = @{IP=19.1.1.1}
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix "$_b;" 					-isVar $false  		-indentlevel 3))
				}
				$PSscriptCode[-1] = $PSscriptCode[-1].TrimEnd() -replace ".$"
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '         };' 				-isVar $false 		-indentlevel 2))
			}
			$PSscriptCode[-1] = $PSscriptCode[-1].TrimEnd() -replace ".$"
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '  }' 							-isVar $False  		-indentlevel 11))

			$ebipaParam			= ' -ebipa $ebipa '

		}

		newLine
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('New-OVLogicalEnclosure -Name "{0}" {1}{2} `' 	-f $name, $enclParam, $egParam ) 	-isVar $false -indent 1) )
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0} `'  										-f  $fwParam ) 						-isVar $false -indent 7) )
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0} `'  										-f  $ebipaParam ) 					-isVar $false -indent 7) )
		newLine


		# --- Scopes
		if ($scopes)
		{
			newLine
			[void]$PSscriptCode.Add( (Generate-PSCustomVarCode -Prefix 'object' -Value 'get-OVLogicalEnclosure | where name -eq $name' -indentlevel 1))
			generate-scopeCode -scopes $scopes -indentlevel 1

		}



		endIfBlock
		# Skip creating because resource already exists
		elseBlock
		[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW {0} already exists.' -f $name ) -isVar $False -indentlevel 1 ))
		endElseBlock

		newLine
	
		# Add disconnect and close the file
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')


		# ---------- Generate script to file
		writeToFile -code $PSscriptCode -file $filename
		
	}
	
	return $ListPS1Files
}

Function Import-YMLlogicalEnclosure([string]$sheetName, [string]$WorkBook, [string]$subdir)
{

	$cSheet, $sheetName 	= $sheetName.Split($SepChar)       # composer	
	$List 					= get-datafromSheet -sheetName $sheetName -workbook $WorkBook  
	$ListYMLfiles	 		= [System.Collections.ArrayList]::new()
	
	foreach ( $le in $List)
	{
		$name          		= $le.name
		$enclosureSN      	= $le.enclosureSerialNumber  	#[]
		$enclosureName   	= $le.enclosureName				#[]
		$enclosureNewName 	= $le.enclosureNewName			#[]
		$enclosureGroup		= $le.enclosureGroup
		$manualAddresses 	= $le.manualAddresses
		$fwBaseline 		= $le.firmwareBaseline
		$fwInstall     		= $le.forceInstallFirmware							# true/false
		$fwvalidateLI     	= $le.validateIfLIFirmwareUpdateIsNonDisruptive     # true/false
		$fwLIupdateMode     = $le.logicalInterconnectUpdateMode					# Parallel/Orchestrated
		$fwUpdateUnmanaged  = $le.updateFirmwareOnUnmanagedInterconnect			# true/false

		# New attribute ?? firmwareUpdateOn: "EnclosureOnly"

		$scopes			    = $le.scopes

		if ($enclosureSN)
		{
			# Create logicalEnclosure filename here per LE
			$filename 			= "$subdir\" + $name.Trim().Replace($Space, '') + '.yml'			
			$YMLscriptCode      = [System.Collections.ArrayList]::new()

			
			[void]$YMLscriptCode.Add((Generate-ymlheader -title 'Configure logical enclosures '))	

			$comment 			= '# ---------- Logical enclosures {0}'	-f $name

			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix $comment -isItem $True											-indentlevel 1 ))

			# --- Get uri of enclosures 
			$snArr 				= $enclosureSN.split($sepChar)
			foreach ($_sn in $snArr)
			{
				$varUri 			= 'var_{0}_uri' 						-f $_sn
				$varName 			= 'var_{0}_Name' 						-f $_sn
				$_filter			= "item.serialNumber == '{0}'"			-f $_sn
				$title 				= 'Get URI for enclosure with SN {0}' 	-f $_sn
				[void]$YMLscriptCode.Add((Generate-ymlTask 	-title $title -isData $False 		-OVTask 'oneview_enclosure_facts'))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix set_fact 			-isVar $True 							-indentlevel 1 ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	$varUri			-value "'{{item.uri}}' "				-indentlevel 2 ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	$varName		-value "'{{item.name}}' "				-indentlevel 2 ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix loop 				-value  "'{{enclosures}}'"				-indentlevel 1 ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix when 				-value  $_filter						-indentlevel 1 ))
			}

			# --- Get uri of enclosure group
			$title 				= 'Get URI for enclosure group {0}' 	-f $enclosureGroup
			[void]$YMLscriptCode.Add((Generate-ymlTask 	-title $title -isData $False 		-OVTask 'oneview_enclosure_group_facts'))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	name			-value $enclosureGroup						-indentlevel 2 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix set_fact 			-isVar $True 								-indentlevel 1 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	var_eg_uri		-value "`"{{enclosure_groups['uri']}}`" "	-indentlevel 2 ))

			# --- Get uri of firmware 
			if ($fwBaseline)
			{
				$title 				= 'Get URI for firmware {0}' 	-f $fwBaseline
				[void]$YMLscriptCode.Add((Generate-ymlTask 	-title $title -isData $False 		-OVTask 'oneview_firmware_driver_facts'))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	name			-value $fwBaseline							-indentlevel 2 ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix set_fact 			-isVar $True 								-indentlevel 1 ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	var_fw_uri		-value "'{{firmware_drivers[0].uri}}' "		-indentlevel 2 ))			
			}
			



			# --- Create logical enclosure 
			$title 				= 'Configure logical enclosures {0}' 	-f $name
			[void]$YMLscriptCode.Add((Generate-ymlTask 	-title $title 	-comment $comment		-OVTask 'oneview_logical_enclosure'))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix name				-value $name						-indentlevel $indentDataStart ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix enclosureGroupUri	-value "'{{var_eg_uri}}'"			-indentlevel $indentDataStart ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix enclosureUris											-indentlevel $indentDataStart ))
			foreach ($_sn in $snArr)
			{
				$varUri 			= 'var_{0}_uri' 		-f $_sn
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix "'{{$varUri}}'"	-isVar $True	-isItem $True		-indentlevel ($indentDataStart +1) ))
			}

			# --- Rename  enclosure
			if ( ($enclosureNewName) -and ($enclosureName -ne $enclosureNewName) ) 
			{
				$snArr 				= $enclosureSN.split($sepChar)
				$nameArr 			= $enclosureName.split($sepChar)
				$newNameArr 		= $enclosureNewName.split($sepChar)

				for ($i=0 ; $i -lt $snArr.Count ; $i++)
				{
					$varName 			= 'var_{0}_Name' 	-f $snArr[$i]
					$newName 			= $newNameArr[$i]

					$title 				= 'Rename enclosures'
					[void]$YMLscriptCode.Add((Generate-ymlTask 	-title $title -iseTag $True			-OVTask 'oneview_enclosure'))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix name				-value "'{{$varName}}'"			-indentlevel $indentDataStart ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix newName				-value $newName					-indentlevel $indentDataStart ))	
				}

			}

			# ------ EBIPA 
			if ($manualAddresses)
			{
				$title 				= 'Configure logical enclosures {0} with EBIPA' 	-f $name
				[void]$YMLscriptCode.Add((Generate-ymlTask 			-title $title 				-state reconfigured		-OVTask 'oneview_logical_enclosure'))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix name				-value $name						-indentlevel $indentDataStart ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix ipAddressingMode	-value 	Manual						-indentlevel $indentDataStart ))

				$frameArr			= $manualAddresses # '\'  -replace $CR, ''				# Remnove extra line
				$frameArr			= $frameArr.Split($Delimiter)		

				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode 	-prefix enclosures											-indentlevel $indentDataStart ))
				for ($i=0; $i -lt $snArr.Count; $i++ )
				{
					$index 				= $i + 1 

					$_fr				= $frameArr | where {$_ -like "*Frame$index=*"} 


					$_frame, $_bay 		= $_fr.Split('@')

					$_bay				= $_bay -replace '{','' -replace '}', ''	
					$_bay 				= $_bay.Replace('Device','|d').replace('Interconnect', '|i')  # Use atrifact to create devivebay and interconnect		
					$_bayArr			= $_bay.Split('|')
					$_devArr 			= $_bayArr | where {$_ -like 'd*'}
					$_icArr 			= $_bayArr | where {$_ -like 'i*'}



					$_sn 				= $snArr[$i]
					$varUri 			= 'var_{0}_uri' 		-f $_sn
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	"'{{$varUri}}'"										-indentlevel ($indentDataStart +1) ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		enclosureUri 		-value "'{{$varUri}}'"		-indentlevel ($indentDataStart +2) ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		deviceBays 										-indentlevel ($indentDataStart +2) ))

					foreach ($_d in $_devArr)
					{
						$_f 			= $_d.Indexof('=') 		# Find first = to get bay number
						$_bayNo			= $_d.Substring(1,($_f-1))	

						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		bayNumber 		-value $_bayNo -isVar $True	-indentlevel ($indentDataStart +3) ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		manualAddresses								-indentlevel ($indentDataStart +3) ))

						$_addr 			= $_d.Substring(($_f + 1))	# Remaining after = --> IPv4Address=10.10.4.13;IPv6Address=0xAA;
						$_addr 			= $_addr.replace($CR, '')
						$_addr			= $_addr.Split(';')
						foreach ($_a in $_addr)
						{
							$_type, $_ip = $_a.Split('=')
							if ($_type)
							{
								$_type 	= $_type.replace('Address', '')
								[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		type		-value $_type 				-indentlevel ($indentDataStart +4) ))
								[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		ipAddress	-value $_ip 				-indentlevel ($indentDataStart +4) ))
							}
						}


					}
					
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		interconnectBays 								-indentlevel ($indentDataStart +2) ))
					foreach ($_d in $_icArr)
					{
						$_f 			= $_d.Indexof('=') 		# Find first = to get bay number
						$_bayNo			= $_d.Substring(1,($_f-1))	

						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		bayNumber 		-value $_bayNo -isVar $True	-indentlevel ($indentDataStart +3) ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		manualAddresses								-indentlevel ($indentDataStart +3) ))

						$_addr 			= $_d.Substring(($_f + 1))	# Remaining after = --> IPv4Address=10.10.4.13;IPv6Address=0xAA;
						$_addr 			= $_addr.replace($CR, '')
						$_addr			= $_addr.Split(';')
						foreach ($_a in $_addr)
						{
							$_type, $_ip = $_a.Split('=')
							if ($_type)
							{
								$_type 	= $_type.replace('Address', '')
								[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		type		-value $_type 				-indentlevel ($indentDataStart +4) ))
								[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		ipAddress	-value $_ip 				-indentlevel ($indentDataStart +4) ))
							}
						}


					}
					
				}







			}

			# ------ Firmware 
			if ($fwBaseline)
			{
				$title 				= 'Update firmware on logical enclosures {0} ' 					-f $name
				[void]$YMLscriptCode.Add((Generate-ymlTask 			-title $title 					-state firmware_updated		-OVTask 'oneview_logical_enclosure'))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix name					-value $name						-indentlevel $indentDataStart ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix firmware													-indentlevel $indentDataStart ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	firmwareBaselineUri 						-value 	"'{{var_fw_uri}}'" -indentlevel ($indentDataStart +1) ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	forceInstallFirmware						-value 	$fwInstall 		-indentlevel ($indentDataStart +1) ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	validateIfLIFirmwareUpdateIsNonDisruptive	-value 	$fwvalidateLI 	-indentlevel ($indentDataStart +1) ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	updateFirmwareOnUnmanagedInterconnect	-value 	$fwUpdateUnmanaged 	-indentlevel ($indentDataStart +1) ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	logicalInterconnectUpdateMode				-value 	$fwLIupdateMode	-indentlevel ($indentDataStart +1) ))


			}

						
		}
		else 
		{
			write-host -ForegroundColor YELLOW "No enclosure serial number. Skip creating logical enclosure...."	
		}






		# ---------- Generate script to file
		YMLwriteToFile -code $YMLscriptCode -file $filename

		[void]$ListYMLfiles.Add($filename)
	}
	return $ListYMLfiles
}


# ---------- Profile and Template are in one function 
Function Import-ProfileorTemplate([string]$sheetName, [string]$WorkBook, [string]$ps1Files, [Boolean]$isSpt, [string]$subdir )
{
    $PSscriptCode             	= [System.Collections.ArrayList]::new()

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
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'class sasJBOD' -isVar $False -indentlevel 0 ))
		startBlock
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '[int]$id'                    -isVar $False -indentlevel 1 ))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '[string]$deviceSlot'         -isVar $False -indentlevel 1 ))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '[string]$name'               -isVar $False -indentlevel 1 ))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '[string]$description'        -isVar $False -indentlevel 1 ))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '[int]$numPhysicalDrives'     -isVar $False -indentlevel 1 ))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '[string]$driveMinSizeGB'     -isVar $False -indentlevel 1 ))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '[string]$driveMaxSizeGB'     -isVar $False -indentlevel 1 ))	
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '[string]$driveTechnology'    -isVar $False -indentlevel 1 ))	
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '[boolean]$eraseData'         -isVar $False -indentlevel 1 ))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '[boolean]$persistent'        -isVar $False -indentlevel 1 ))
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '[string]$sasLogicalJBODUri'  -isVar $False -indentlevel 1 ))
		endBlock
	}


	connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode
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



		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('write-host -foreground CYAN "----- Creating profile {0} "' -f $spName) -isVar $False ))
 
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

		$value 					= $getCmd + " -name '{0}' -ErrorAction SilentlyContinue " -f $spName #HKD
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'profile' 						-Value $value ))		
		
		ifBlock			-condition 'if ($Null -eq $profile )' 
		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('# -------------- Attributes for profile {0} ' -f $name) -isVar $False -indentlevel 1))

		[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix name -value ("'{0}'" -f $spName) -indentlevel 1 ))
		
		$descParam 					=  $null
		if ($description) 
		{
			$descParam 				=  ' -description $description'
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix description -value ("'{0}'" -f $description) -indentlevel 1 ))
		}

		# ############### Server Profile Region
		$spDescParam 				=  $null
		if ( $isSpt -and $spDescription)
		{
			$spdescParam 			= ' -ServerProfileDescription $spDescription'
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix spDescription -value ("'{0}'" -f $spDescription) -indentlevel 1 ))
		}

		# --- server profile template
		# Used when creating profile from template
		$spTemplateParam 	= $null
		if ($template)
		{
			$spTemplateParam = ' -ServerProfileTemplate $spTemplate'
			$value 			 = "Get-OVServerProfileTemplate -name '{0}'" -f $template
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix spTemplate -value $value -indentlevel 1 ))
		}
		else # Standalone profile
		{
			# -- server hardware type	
			$shtParam 			= ' -ServerHardwareType $sht'
			$value 				= "Get-OVserverHardwareType -name '{0}'" -f $sht

			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix sht -value $value -indentlevel 1 ))

			# -- enclosure group
			$egParam 			= ' -EnclosureGroup $eg'
			$value 				= "Get-OVEnclosureGroup -name '{0}'" -f $eg
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix eg -value $value -indentlevel 1 ))

			# --- affinity
			$affinityParam 		= ' -affinity $affinity'
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix affinity -value ("'{0}'" -f $affinity) -indentlevel 1 ))

		}

		# ---- server hardware
		$hwParam 			= $null
		if ($serverHardware)
		{
			$hwParam 		= ' -Server $server'
			$value 			 = "Get-OVServer -name '{0}'" -f $serverHardware
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix server -value $value -indentlevel 1 ))
			# -- Add code to power off server
			$value 			= 'Stop-OVServer -inputObject $server -force -Confirm:$False| Wait-OVTaskComplete	'
			[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix status -value $value  -indentlevel 1))
				
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
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '# -------------- Firmware Baseline section ' -isVar $False -indentlevel 1))
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'manageFirmware' -Value ('${0}' -f $manageFirmware) -indentlevel 1))
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'sppName' -Value ('"{0}"' -f $fwBaseline) -indentlevel 1))
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'fwBaseline' -Value 'Get-OVbaseline -SPPname $sppName' -indentlevel 1))
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'fwInstallType' -Value ('"{0}"' -f $fwInstallType) -indentlevel 1))
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'fwForceInstall' -Value ('${0}' -f $fwforceInstall)-indentlevel 1))
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'fwActivation' -Value ('"{0}"' -f $fwActivation) -indentlevel 1))
				
				$fwParam = ' -firmware -Baseline $fwBaseline -FirmwareInstallMode $fwInstallType -ForceInstallFirmware:$fwForceInstall -FirmwareActivationMode $fwActivation '

				$fwConsistencyParam	= $fwScheduleParam = $null
				if ($isSpt)
				{
					if ($null -eq $fwConsistency)
					{
						$fwConsistency 	= 'None'
					}
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'fwConsistency' -Value ('"{0}"' -f $fwConsistency) -indentlevel 1))
					$fwConsistencyParam	= ' -FirmwareConsistencyChecking $fwConsistency'
				}
				else   # SP specific here
				{
					if ($fwSchedule)
					{
						[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'fwSchedule' -Value ('[DateTime]{0}' -f $fwSchedule) -indentlevel 1))
						$fwScheduleParam = ' -FirmwareActivateDateTime $fwSchedule'
					}
				}

				$fwParam += $fwConsistencyParam	+ $fwScheduleParam 
			}


			# ############### Connections
			newLine
			[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '# -------------- Connections section ' -isVar $False -indentlevel 1))

			$connectionsParam 			= $null
			if ($manageConnections)
			{
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'manageConnections' -Value ('${0}' -f $manageConnections) -indentlevel 1))
				$connectionsParam  	 	= ' -manageConnections $manageConnections'
			
				if ($isSpt)
				{
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'connectionsConsistency' -Value ('"{0}"' -f $connConsistency) -indentlevel 1))
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
					$targetLUN 			= $conn.targetLun
					$bootTarget 		= $conn.bootTarget


					if ($isSp)
					{
						### Custom MAC and WWPN
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
							$_wwpnParam = if ($wwpn)	{ ' -wwpn {0} ' -f $wwpn}	else {''}

							#HKD01
							$userDefinedParam = ' -UserDefined:$True ' + $_macParam + $_wwpnParam + $_wwnnParam
						}

						### Boot from SAN
						$_bootFromSAN 		= " -bootVolumeSource $bootVolumeSourcc "
						if ($bootVolumeSource -eq 'UserDefined')
						{
							$_bootFromSAN 	+= " -TargetWwpn $bootTarget -LUN $targetLUN " #HKD02
						}

					}


					$value 			= "Get-OVnetwork -name '{0}' -ErrorAction SilentlyContinue " -f $network		#HKD
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'network' 						-Value $value -indentlevel 1))		
					ifBlock -condition 'if ($null -eq $network)'  -indentlevel 1
					$value 			= "Get-OVnetworkSet -name '{0}' -ErrorAction SilentlyContinue " -f $network		#HKD
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'network' 						-Value $value -indentlevel 2))
					endIfBlock -indentLevel 1



					if ($bootable)
					{
						$bootParam 	= ' -bootable:${0} -priority {1} {2} ' -f $bootable,$priority,$_bootFromSAN
					}

					# TBD FibreChannel Bfs

					$value 				+= ' -network $network'
					$value 				+= $bootParam

					$nameParam 			= if ($name) {' -name "{0}" ' -f $name} else {''}

					# -- code
					$_connection		= '$' + "conn$index"
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ( '{0}   = New-OVServerProfileConnection {1} -ConnectionID {2} `' 	-f $_connection,$nameParam, $id)  		-isVar $False  -indentlevel 1))
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ( '  -PortId "{0}" -RequestedBW {1} `' 								-f $portId , $requestedMbps) 		-isVar $False  -indentlevel 11))
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ( '  -network $network {0} {1} ' 										-f $bootParam, $userDefinedParam)	-isVar $False  -indentlevel 11))
					
					newLine
					[void]$connectionArray.Add($_connection)
					$index++

				}

				if ($connectionArray)
				{
					$value 					= '@(' + ($connectionArray -join $comma) + ')'
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'connectionList' -Value $value -indentlevel 1))
					$connectionsParam 		= ' -Connections $connectionList '
					$connectionArray       	= [System.Collections.ArrayList]::new()
				}
			


			# ############### Local Storage

			$localStorageParam				= $null
			$lsList 						= $localStorageList  | where ProfileName -eq $spName

			if ($Null -ne $lsList)
			{
				newLine
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '# -------------- Local Storage section ' -isVar $False -indentlevel 1))
				if ($isSpt)
				{
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'lsConsistencyChecking' -Value ('"{0}"' -f $lsConsistency) -indentlevel 1))
					$localStorageParam		= ' -LocalStorageConsistencyChecking $lsConsistencyChecking'
				}
				newLine
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '# --- Search for SAS-Logical-Interconnect for SASlogicalJBOD ' -isVar $False -indentlevel 1))
				
				# Find SAS-Logical-INTERCONNECT from logical enclosure
				# Note: $eg is defined earlier in the generated script

				# Step 1 - find SasLIG from EnclosureGroup
				$value 				= 'Search-OVAssociations ENCLOSURE_GROUP_TO_LOGICAL_INTERCONNECT_GROUP -parent $eg' 
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ligAssociation -value $value -indentlevel 1 ))

				$value 				= '($ligAssociation | where ChildUri -like "*sas*").ChildUri'
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix sasLigUri -value $value -indentlevel 1 ))

				$value 				= 'if ($sasLigUri) { Send-OVRequest -uri $sasLigUri } else {$Null}'
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix sasLig -value $value -indentlevel 1 ))

				# Step 2 - find Sas Interconnect from SAS Lig
				$value 				= '(Search-OVAssociations LOGICAL_INTERCONNECT_GROUP_TO_LOGICAL_INTERCONNECT -Parent $sasLig).childUri[0]' 
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix sasUri -value $value -indentlevel 1 ))

				$value 				= 'if ($sasUri) { Send-OVRequest -uri $sasUri } else {$Null}'
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix sasLI -value $value -indentlevel 1 ))

				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix lsJBOD -value '[System.Collections.ArrayList]::new()' -indentlevel 1 ))
			
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
								[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix $prefix	-isVar $False 	-indentlevel 1))
								
								$value 			= "'{0}'" -f $_ld
								[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'ldName'			-value $value 	-indentlevel 1))
								$ldParam 		+= ' -Name $ldName'
								
								$this_index 	= [array]::IndexOf($logicalDrivesArray, $_ld )


								$value 			= '${0}' -f $bootableArray[$this_index] 
								[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'bootable'		-value $value 	-indentlevel 1))
								$ldParam 		+= ' -Bootable $bootable'
								
								$value 			= "'{0}'" -f $raidLevelArray[$this_index]	
								[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'raidLevel'		-value $value 	-indentlevel 1))
								$ldParam 		+= ' -RAID $raidLevel'

								$_driveType 	= $driveTypeArray[$this_index]
								if ($_driveType)
								{
									$value 			= "'{0}'" -f $_driveType	
									[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'driveType'		-value $value 	-indentlevel 1))	
									$ldParam 		+= ' -DriveType $driveType' 
								}

								$value 			= "{0}" -f $physDrivesArray[$this_index]
								[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'numberofDrives'	-value $value 	-indentlevel 1))	
								$ldParam 		+= ' -NumberofDrives $numberofDrives '


								$value 			= "'{0}'" -f $acceleratorArray[$this_index]	
								[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'accelerator'		-value $value 	-indentlevel 1))
								$ldParam 		+=' -Accelerator $accelerator'



								# Make sure that there is no space after backstick (`)
								$logicalDisk 	= '{0}' -f "LogicalDisk$diskIndex"
								[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix $logicalDisk  -value ('New-OVServerProfileLogicalDisk {0} ' -f $ldParam)	-indentlevel 1))

								[void]$logicaldiskArray.Add('${0}' -f $logicalDisk)
								$diskIndex++
							}
						}
					
				
						
						# --- Generate array of logical disk for this controller
						if ($logicaldiskArray)
						{
							$logicalDisks 		= '@(' + ($logicaldiskArray -join $comma) + ')'
							[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'logicalDisks'  -value $logicalDisks -indentlevel 1))
							$logicalDiskParam 	= ' -LogicalDisk $LogicalDisks '
						}
						else 
						{
							$logicalDiskParam 	= ''
						}
			

						$controllerParam  		= ' -ControllerID "{0}" -Mode "{1}" -Initialize:${2} -WriteCache "{3}" {4}' -f $deviceSlot, $mode, $initialize, $writeCache, $logicalDiskParam 
						$controller				= '{0}' -f "controller$contIndex"  

						# ---- Generate new Disk Controller
						[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix $controller  -value ('New-OVServerProfileLogicalDiskController {0}' -f $controllerParam) -indentlevel 1))
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
							[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix availableDrives  -value $value -indentlevel 2))
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
								[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix $jbodDisk  -value $value -isVar $False -indentlevel 3))
								$jbodIndex++
								ifBlock -condition "if ($jbodDisk)"  -indentlevel 3
									[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '_jbod' 					-value 'new-object -type sasJBOD' 	-indentlevel 4))
									[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '_jbod.id'                -value $driveID                     -indentlevel 4 ))
									[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '_jbod.deviceSlot'        -value $deviceSlot                  -indentlevel 4 ))
									[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '_jbod.name'              -value $logicalDrives               -indentlevel 4 ))
									if ($driveDescription)
									{
										[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '_jbod.description'   -value $driveDescription            -indentlevel 4 ))
									}
									[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '_jbod.numPhysicalDrives' -value $numPhysicalDrives           -indentlevel 4 ))
									[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '_jbod.driveMinSizeGB'    -value $driveMinSize                -indentlevel 4 ))
									[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '_jbod.driveMaxSizeGB'    -value $driveMaxSize                -indentlevel 4 ))	
									[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '_jbod.driveTechnology'   -value $driveTechnology             -indentlevel 4 ))	
									[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '_jbod.eraseData'         -value $eraseData                   -indentlevel 4 ))
									[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '_jbod.persistent'        -value $persistent                  -indentlevel 4 ))
									[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '_jbod.sasLogicalJBODUri' -value ('{0}.uri' -f $jbodDisk)     -indentlevel 4 ))

									[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '[void]$lsJBOD.Add($_jbod)'  -isVar $False 					-indentlevel 4 ))

								endIfBlock -indentlevel 3							
							endIfBlock -indentlevel 2
						endIfBlock -indentlevel 1

					}

				}

				if ($controllerArray)
				{
					# ----- Generate params for profiles
					$controllers 	= '@(' + ($controllerArray -join $comma) + ')'
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'controllers'  -value $controllers  -indentlevel 1))
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
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '# -------------- SAN Storage section ' -isVar $False -indentlevel 1))
				if ($isSpt)
				{
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'sanConsistencyChecking' -Value ('"{0}"' -f $sanConsistency) -indentlevel 1))
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
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix "$_volVariable	=	@{" 	-isVar $False 	-indentlevel 1))
					if ($_lunType -eq 'Manual')
					{
						[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix "id						= $_lun $SepHash"  		-isVar $False 	-indentlevel 6))
					}
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix "lunType					= '$_lunType' $SepHash"	-isVar $False 	-indentlevel 6))
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix "volumeUri				= $_volUri $SepHash" 	-isVar $False 	-indentlevel 6))
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix "volumeStorageSystemUri	= $_stsUri $SepHash" 	-isVar $False 	-indentlevel 6))			
					if ($_stPath)
					{
						[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix "storagePaths			= $_stPath " 			-isVar $False 	-indentlevel 6))
					}
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode 	-Prefix '}'												-isVar $False 	-indentlevel 6))

					[void]$storageVolumeArray.add($_volVariable)

				}
				if ($storageVolumeArray )
				{
					$value 						= "@( " + ($storageVolumeArray -join $Comma) + " )"
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix storageVolume  	-value $value	-indentlevel 1))
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
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '# -------------- Boot mode section ' -isVar $False -indentlevel 1))
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'bootMode' 		-Value ('"{0}"' -f $bmMode) -indentlevel 1))
				

				$bmParam 		= ' -bootMode $bootMode '
				if ($pxeBootPolicy)
				{
					$bmParam 	+=  ' -PxeBootPolicy {0}' -f  $pxeBootPolicy
				}

				if ($bootMode -match 'UEFI Optimized*')
				{
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'secureBoot' 		-Value ('"{0}"' -f $secureBoot) -indentlevel 1))
					$bmParam 	+= ' -SecureBoot $secureBoot'
				}

				if ($order)
				{
					$bootOrder 	 = "@('" + $order.Replace($SepChar,"'$Comma'") + "')" 
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'bootOrder' 		-Value ('{0}' -f $bootOrder) -indentlevel 1)) 
					$boParam     = ' -ManageBoot ${0} -BootOrder $bootOrder' -f $bo
				}

				$bmConsistencyParam	= $null
				$boConsistencyParam = $null

				if ($isSpt)
				{
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'bmConsistency' -Value ('"{0}"' -f $bmConsistency) -indentlevel 1))
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'boConsistency' -Value ('"{0}"' -f $boConsistency) -indentlevel 1))
					
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
				if ($biosSettings)
				{ 
					$settingsArr 	= $biosSettings.Split($SepChar)
				}

				newLine
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '# -------------- BIOS section ' -isVar $False -indentlevel 1))
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'biosSettings' 	-Value '@(' -indentlevel 1))

				foreach ($setting in $settingsArr)
				{
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ("{0}$COMMA" -f $setting)	-isVar $false  -indentlevel 2))
				}
				$PSscriptCode[-1] = $PSscriptCode[-1] -replace $COMMA , ''
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ')'	-isVar $false  -indentlevel 2))



				$biosParam 		= ' -Bios -BiosSettings $biosSettings'

				if ($isSpt)
				{
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'biosConsistency' -Value ('"{0}"' -f $biosConsistency) -indentlevel 1))				
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
				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix '# -------------- iLO section ' -isVar $False -indentlevel 1))

				[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'manageIlo' -Value ('${0}' -f $manageIlo) -indentlevel 1))
				$iloParam  	 			= ' -ManageIloSettings $manageIlo'
				if ($isSpt)
				{
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix 'iloConsistency' -Value ('"{0}"' -f $iloConsistency) -indentlevel 1))				
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
											$_privParamArr	+= $iLOPrivilegeParamEnum.Item($priv)
										}
									}

								
									$iloUser 			= '$iloUser{0}' -f $iloUserIndex++
									$value 				= 'new-OVIloLocalUserAccount	{0}{1} `' 		-f $iloNameParam, $iloDisplayParam 
									[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix $iloUser 	-value $value 			-isVar $False	-indentlevel 1)) 
									[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0} `'		-f $iloPasswordParam)	-isVar $False	-indentlevel 18))
									foreach ($_param in $_privParamArr)
									{
										[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0} `'	-f $_param)		-isVar $False	-indentlevel 18))
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
											$_privParamArr			+= $iLOPrivilegeParamEnum.Item($priv)
										}
									}

									$iloDirectoryGroup 			= '$iloDG{0}' -f $ilodirectoryGroupIndex++
									[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix $iloDirectoryGroup 	-value ('New-OVIloDirectoryGroup{0} `' -f $iloDirGroupParam1) -isVar $False -indentlevel 1))
									foreach ($_param in $_privParamArr)
									{
										[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0} `'	-f $_param)		-isVar $False	-indentlevel 18))
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
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode 	-Prefix 'iloLocalAccounts' -value ('@({0})' -f ($iloLocalAccounts -join $COMMA))	-indentlevel 1))
				}
				
				if ($iloDirectoryGroups)
				{
					$iloDirectoryGroupsParam 	= ' -ManageDirectoryGroups -DirectoryGroups $iloDirectoryGroups '
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode 	-Prefix 'iloDirectoryGroups' -value ('@({0})' -f ($iloDirectoryGroups -join $COMMA))	-indentlevel 1))
				}

				[void]$PSscriptCode.Add((Generate-PSCustomVarCode 	-Prefix 'iloPolicy' -value ('new-OVServerProfileIloPolicy	{0} `' -f $iloAdminParam)	-indentlevel 1))

				
				if ($iloLocalAccountsParam)
				{
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0} `' -f $iloLocalAccountsParam) -isVar $False	-indentlevel 19))
				}
				if ($iloDirectoryGroupsParam)
				{
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0} `' -f $iloDirectoryGroupsParam) -isVar $False	-indentlevel 19))
				}

				if ($iloHostNameParam)
				{
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0} `' -f $iloHostNameParam) 	-isVar $False	-indentlevel 19))
				}
				foreach ($_param in $_dirParamArr)
				{
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0} `' -f $_param) -isVar $False	-indentlevel 19))
				}
				foreach ($_param in $_kmParamArr)
				{
					[void]$PSscriptCode.Add((Generate-PSCustomVarCode -Prefix ('{0} `' -f $_param) -isVar $False	-indentlevel 19))
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
			[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix $prefix -isVar $False -indentlevel 1))
			newLine   # to end the command
		}
		else # SPT or standalone profile
		{
			$prefix 	= $newCmd + '	 	 -Name $name {0}{1}{2}{3}{4} `' -f $descParam, $spDescParam, $shtParam, $egParam, $affinityParam
			[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix $prefix -isVar $False -indentlevel 1))

			if ($fwParam)
			{
				$prefix		= '{0} `' -f $fwParam
				[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix $prefix -isVar $False -indentlevel 9))
			}
			if ($bmParam)
			{
				$prefix		= '{0} `' 	-f $bmParam
				[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix $prefix -isVar $False -indentlevel 9))
			}
			if ($boParam)
			{
				$prefix		= '{0} `' 	-f $boParam
				[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix $prefix -isVar $False -indentlevel 9))
			}
			if ($biosParam)
			{
				$prefix		= '{0} `' 	-f $biosParam
				[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix $prefix -isVar $False -indentlevel 9))
			}

			if ($connectionsParam)
			{
				$prefix		= '{0} `' 	-f $connectionsParam
				[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix $prefix -isVar $False -indentlevel 9))
			}

			if ($localStorageParam)
			{
				$prefix		= '{0} `' 	-f $localStorageParam
				[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix $prefix -isVar $False -indentlevel 9))
			}

			if ($StorageVolumeParam)
			{
				$prefix		= '{0} `' 	-f $StorageVolumeParam
				[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix $prefix -isVar $False -indentlevel 9))
			}

			if ($iloParam)
			{
				$prefix		= '{0}' 	-f $iloParam
				[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix $prefix -isVar $False -indentlevel 9))
			}

			newLine   # to end the command



			# sasLogicalJBOD

			ifBlock -condition 'if ($lsJBOD)' -indentlevel 1
				[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix '# ------ Configure saslogicalJBOD for profile' -isVar $False 	-indentlevel 2 ))
				[void]$PSscriptCode.Add( (Generate-PSCustomVarCode -Prefix 'prf' 		-Value ($getCmd + ' | where name -eq $name')     -indentlevel 2))
				ifBlock -condition 'if ($prf)' -indentlevel 2
					[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix '$lsJBOD | % { $prf.localStorage.sasLogicalJBODs += $_ }' -isVar $False 	-indentlevel 3 ))		
				endIfBlock -indentlevel 2
				[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('{0} -InputObject $prf' -f $saveCmd) -isVar $False 	-indentlevel 2 ))
			endIfBlock -indentlevel 1

		}
		# --- Scopes
		if ($scopes)
		{
			newLine
			[void]$PSscriptCode.Add( (Generate-PSCustomVarCode -Prefix 'object' -Value ($getCmd + ' | where name -eq $name') -indentlevel 1))
			generate-scopeCode -scopes $scopes -indentlevel 1

		}


		endIfBlock
		# Skip creating because resource already exists
		elseBlock
		[void]$PSscriptCode.Add(( Generate-PSCustomVarCode -Prefix ('write-host -foreground YELLOW "{0} already exists." ' -f $spName ) -isVar $False -indentlevel 1 ))
		endElseBlock

		newLine

		# ---------------- Split in different scripts for server profiles ONLY!
		if ($isSP)
		{
			if ($i -ge $MAXPROFILE)
			{
				newLine
	
				# Add disconnect and close the file
				[void]$PSscriptCode.Add('Disconnect-OVMgmt')
			
				# ---------- Generate script to file
				writeToFile -code $PSscriptCode -file $spfilename
				[void]$ListPS1Files.Add($spfilename)
	
				# Build a new set
				$PSscriptCode             	= [System.Collections.ArrayList]::new()
				$count++
				$spfilename 				= "$subdir\profileGroup$count" + ".PS1"
				connect-Composer  -sheetName $cSheet    -workBook $WorkBook -PSscriptCode $PSscriptCode

				$i = 1
			}
			
		}

	}

	if ($spList)
	{
		[void]$PSscriptCode.Add('Disconnect-OVMgmt')
	}




	if ($isSp)
	{
		$ps1Files 	= $spFilename
		[void]$ListPs1Files.Add($spfilename)
	}


	 # ---------- Generate script to file
	 writeToFile -code $PSscriptCode -file $ps1Files

	 return $ListPS1Files

}

Function Import-YMLProfileorTemplate([string]$sheetName, [string]$WorkBook, [string]$YMLFiles, [Boolean]$isSpt, [string]$subdir )
{
	$YMLscriptCode         = [System.Collections.ArrayList]::new()
	

	$cSheet,$spSheet,$connSheet,$localStorageSheet,$SanStorageSheet,$iLOSheet 		= $sheetName.Split($SepChar)

	$spList 					= if ($spSheet)	 			{get-datafromSheet -sheetName $spSheet -workbook $WorkBook				} else {$null}
	$connList 					= if ($connSheet) 			{get-datafromSheet -sheetName $connSheet -workbook $WorkBook 			} else {$null}
	$localStorageList 			= if ($localStorageSheet)	{get-datafromSheet -sheetName $localStorageSheet -workbook $WorkBook 	} else {$null}	
	$sanStorageList 			= if ($sanStorageSheet) 	{get-datafromSheet -sheetName $sanStorageSheet -workbook $WorkBook		} else {$null}
	$iLOList 					= if ($iLOSheet) 			{get-datafromSheet -sheetName $iLOSheet -workbook $WorkBook				} else {$null}


	$isSP 						= -not $isSpt



	[void]$YMLscriptCode.Add((Generate-ymlheader -title 'Configure server profiles or templates'))



	foreach ( $prof in $spList)
	{

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

		$macType 				= $prof.macType
		$wwnType 				= $prof.wwnType
		$snType 				= $prof.serialNumberType
		$iscsiType				= $prof.iscsiInitiatorNameType	
		$hideFlexNics 			= $prof.hideUnusedFlexNics


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

		$getCmd		= $newCmd =  $saveCmd = $null
		if ($isSpt)
		{
			$getFact 			= 'oneview_server_profile_template_facts'
			$newFact 			= 'oneview_server_profile_template'
			$saveFact 			= '*****oneview_server_profile_template_facts'
			$_spType 			= $YMLtype540Enum.item('serverProfileTemplate')
			$comment 			= '# ---------- Server Profile Template {0}' 	-f $spName
			$_task 				= ' Create server profile template  {0}' 		-f $spName
		}
		else 
		{
			$getFact 			= 'oneview_server_profile_facts'
			$newFact 			= 'oneview_server_profile'
			$saveFact 			= '*****oneview_server_profile_template_facts'
			$_spType 			= $YMLtype540Enum.item('serverProfile')
			$comment 			= '# ---------- Server Profile {0}' 			-f $spName
			$_task 				= ' Create server profile {0}' 					-f $spName
		}

		newLine	-code $YMLscriptCode
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix $comment  				-isItem $True  								-indentlevel 2 ))

		# --- Power Off server
		if (($isSp) -and ($null -ne $serverHardware))
		{
			newLine	-code $YMLscriptCode
			$title 			= ' Power off server {0}' 	-f $serverHardware
			[void]$YMLscriptCode.Add((Generate-ymlTask 		 	-title $title  -OVTask 'oneview_server_hardware'))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		name				-value "'$serverHardware'"				-indentlevel $indentDataStart )) 
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		powerStateData 												-indentlevel $indentDataStart ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			powerState			-value "'Off'" 						-indentlevel ($indentDataStart +1) ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			powerControl		-value "'MomentaryPress'" 			-indentlevel ($indentDataStart +1) ))
		}

		# New Server profile or template

		$title 						= $_task
		if ($Template)
		{
			$title 			= ' Create server profile {0} from template  {1}' 		-f $spName, $template
		}

		[void]$YMLscriptCode.Add((Generate-ymlTask 		 	-title $title -comment $comment -OVTask $newFact))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		name				-value $spName							-indentlevel $indentDataStart ))
		[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		type				-value $_spType							-indentlevel $indentDataStart ))	

		if ($description)
		{
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	description			-value $description						-indentlevel $indentDataStart ))	
		}

		if ($isSpt -and $spDescription)
		{
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	serverProfileDescription -value $spDescription				-indentlevel $indentDataStart ))	
		}
		
		if ($template)
		{
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	serverProfileTemplateName	-value $template				-indentlevel $indentDataStart ))	
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	serverHardwareName			-value $serverHardware			-indentlevel $indentDataStart ))
		}
		else # either SPT or standalone SP
		{
			# -- server hardware type	
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	serverHardwareTypeName		-value $sht						-indentlevel $indentDataStart ))

			# -- enclosure group	
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	enclosureGroupName			-value $eg						-indentlevel $indentDataStart ))

			# -- affinity	
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	affinity					-value $affinity				-indentlevel $indentDataStart ))

			# -- server hardware
			if ($isSp)
			{
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix serverHardwareName			-value $serverHardware			-indentlevel $indentDataStart ))				
			}

			# ############### Firmware Region
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

				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix firmware														-indentlevel $indentDataStart ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	manageFirmware			-value $manageFirmware				-indentlevel ($indentDataStart + 1) ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	firmwareBaselineName	-value $fwBaseline					-indentlevel ($indentDataStart + 1) ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	firmwareInstallType		-value $fwInstallType				-indentlevel ($indentDataStart + 1) ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	forceInstallFirmware	-value $fwforceInstall				-indentlevel ($indentDataStart + 1) ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	firmwareActivationType	-value $fwActivation				-indentlevel ($indentDataStart + 1) ))
				if ($isSpt)
				{
					$_fwConsistency 		= $YMLconsistencyCheckingSptEnum.item($fwConsistency)
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix complianceControl		-value $_fwConsistency				-indentlevel ($indentDataStart + 1) ))
				}
				## HKD TO BE ADDED
				# 	$fwSchedule
			}

			# ############### Boot Settings : BootMode / BootOrder
			if ($bm )
			{
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix bootMode														-indentlevel $indentDataStart ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	manageMode			-value $bm								-indentlevel ($indentDataStart + 1) ))	
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	mode				-value $bmMode							-indentlevel ($indentDataStart + 1) ))
				if ($isSpt)
				{
					$_bmConsistency 		= $YMLconsistencyCheckingSptEnum.item($bmConsistency)
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix complianceControl	-value $_bmConsistency					-indentlevel ($indentDataStart + 1) ))
				}
				
				# For Ansible, 'Auto '= null

				if (($pxeBootPolicy) -and ($pxeBootPolicy -notlike 'Auto'))
				{
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix pxeBootPolicy		-value $pxeBootPolicy					-indentlevel ($indentDataStart + 1) ))		
				}
			}
			if ($order )
			{
				$_orderArray = $order.split($SepChar) 

				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix boot															-indentlevel $indentDataStart ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	manageBoot			-value $bo								-indentlevel ($indentDataStart + 1) ))
				if ($isSpt)
				{
					$_boConsistency 		= $YMLconsistencyCheckingSptEnum.item($boConsistency)
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix complianceControl	-value $_boConsistency					-indentlevel ($indentDataStart + 1) ))
				}
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	order				-value '['								-indentlevel ($indentDataStart + 1) ))
				foreach ($_order in $_orderArray)
				{
					$_item 					= '{0} ,' -f $_order
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	$_item	 		-isItem $True							-indentlevel ($indentDataStart + 2) ))
				}
				$YMLscriptCode[-1] 			= $YMLscriptCode[-1].TrimEnd() -replace ".$"
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	' '					-value ' ]'		-isItem $True			-indentlevel ($indentDataStart + 1) ))
				
			}


			# ############### BIOS Settings
			if ($bios)
			{
				if ($biosSettings)
				{
					$biosSettingsArr 	= $biosSettings.Replace("@{ ","").Replace("}","").Split($SepChar)
				}

				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix bios															-indentlevel $indentDataStart ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	manageBios			-value $bios							-indentlevel ($indentDataStart + 1) ))	
				if ($isSpt)
				{
					$_biosConsistency 		= $YMLconsistencyCheckingSptEnum.item($biosConsistency)
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix complianceControl	-value $_biosConsistency				-indentlevel ($indentDataStart + 1) ))
				}
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	overriddenSettings	-value '['								-indentlevel ($indentDataStart + 1) ))	
				foreach ($_bios in $biosSettingsArr)
				{
					$_id, $_value 	= $_bios.Split(';').Trim()
					$_id 			= $_id.Split('=').Trim()				# id = "Sriov"
					$_value 		= $_value.Split('=').Trim()				# value = "Enabled"

					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'{'										-isItem $True	-indentlevel ($indentDataStart + 2)  ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		"id"			-value ($_id[1] + $COMMA)			-indentlevel ($indentDataStart + 3)  ))	
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		"value"			-value $_value[1]					-indentlevel ($indentDataStart + 3)  ))	
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'},'									-isItem $True	-indentlevel ($indentDataStart + 2)  ))
				}
				$YMLscriptCode[-1]	= $YMLscriptCode[-1].TrimEnd() -replace ".$"
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	' '					-value ' ]'		-isItem $True			-indentlevel ($indentDataStart + 1) ))


			}

			# ############### Connections
			if ($manageConnections)
			{
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix connectionSettings												-indentlevel $indentDataStart  ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	manageConnections			-value $manageConnections		-indentlevel ($indentDataStart + 1) ))	
				if ($isSpt)
				{
					$_connConsistency 		= $YMLconsistencyCheckingSptEnum.item($connConsistency)
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix complianceControl	-value $_connConsistency					-indentlevel ($indentDataStart + 1) ))
				}	

				$connectionArray       	= [System.Collections.ArrayList]::new()
				$connectionList 		= $connList | where ProfileName -eq $spName
				$index					= 1

				if ($connectionList)
				{
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix connections													-indentlevel ($indentDataStart + 1) ))
				}
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


					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix "- id"					-value $id						-indentlevel ($indentDataStart + 1)  ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix "  portId"				-value $portId					-indentlevel ($indentDataStart + 1)  ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix "  name"				-value $name					-indentlevel ($indentDataStart + 1)  ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix "  functionType"		-value $functionType			-indentlevel ($indentDataStart + 1)  ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix "  networkName"			-value $network					-indentlevel ($indentDataStart + 1)  ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix "  requestedMbps"		-value $requestedMbps			-indentlevel ($indentDataStart + 1)  ))
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix "  requestedVFs"		-value $requestedVFs			-indentlevel ($indentDataStart + 1)  ))
					
					if ($lagName)
					{
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix "  lagName"			-value $lagName					-indentlevel ($indentDataStart + 1)  ))
					}
					
					if ($bootable)
					{
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix "  boot"												-indentlevel ($indentDataStart + 1)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	"  priority"			-value $priority			-indentlevel ($indentDataStart + 2)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	"  bootVolumeSource"	-value $bootVolumeSource	-indentlevel ($indentDataStart + 2)  ))
					}

					if ( ($isSp) -and ($conn.userDefined) )
					{
						$mac	 	= $Conn.mac
						$wwpn		= $Conn.wwpn
						$wwnn		= $Conn.wwnn

						if ($mac)
						{
							[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix "  mac"				-value $mac					-indentlevel ($indentDataStart + 1)  ))
						}
						if ($wwpn)
						{
							[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix "  wwpn"			-value $wwpn					-indentlevel ($indentDataStart + 1)  ))
						}			
						if ($wwnn)
						{
							[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix "  wwnn"			-value $wwnn					-indentlevel ($indentDataStart + 1)  ))
						}
					}				
				
				}
			}


			# ############### iLO Settings
			if ($manageIlo)
			{
				$iLOsettingList 		= $iLOList | where ProfileName -match $spName

				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	managementProcessor											-indentlevel $indentDataStart ))
				[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		manageMp			-value $manageIlo					-indentlevel ($indentDataStart + 1) ))	
				if ($isSpt)
				{
					$_iloConsistency 		= $YMLconsistencyCheckingSptEnum.item($iloConsistency)
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	complianceControl	-value $_iloConsistency				-indentlevel ($indentDataStart + 1) ))

				}

				if ($iLOsettingList)
				{
					[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	mpSettings												-indentlevel ($indentDataStart + 1) ))
				

					$_adminAccountArr 		= $iLOsettingList | where settingType -eq 'AdministratorAccount'
					$_hostNameArr 			= $iLOsettingList | where settingType -eq 'HostName'
					$_localAccountArr 		= $iLOsettingList | where settingType -eq 'LocalAccounts'
					$_dirGroupsArr 			= $iLOsettingList | where settingType -eq 'DirectoryGroups'
					$_KeyManagerArr 		= $iLOsettingList | where settingType -eq 'KeyManager'
					$_directoryArr 			= $iLOsettingList | where settingType -eq 'Directory'

					#------- AdministratorAccount
					foreach ($s in $_adminAccountArr)
					{

						$_deleteAdmin 			= 'deleteAdministratorAccount : {0},'			-f $s.deleteAdministratorAccount 
						$_password 				= "password:                   '{0}'"			-f $s.adminPassword

						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'- {'										  	-isItem $True	-indentlevel ($indentDataStart + 2)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			'settingType : AdministratorAccount,' 	-isItem $True	-indentlevel ($indentDataStart + 4)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			'args        : { '					   	-isItem $True 	-indentlevel ($indentDataStart + 4) ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 				$_deleteAdmin					   	-isItem $True 	-indentlevel ($indentDataStart + 5)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 				$_password							-isItem $True 	-indentlevel ($indentDataStart + 5) ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			'              } '						-isItem $True	-indentlevel ($indentDataStart + 4) ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'  } '											-isItem $True	-indentlevel ($indentDataStart + 2)  ))
					}

					#------- HostName
					foreach ($s in $_hostNameArr)
					{
						$_hostName 				= "hostName: '{0}'"						-f $s.hostName 
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'- {'										-isItem $True	-indentlevel ($indentDataStart + 2)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'settingType : Hostname, '				-isItem $True	-indentlevel ($indentDataStart + 4) ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'args        : { '						-isItem $True	-indentlevel ($indentDataStart + 4) ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			$_hostName 							-isItem $True	-indentlevel ($indentDataStart + 5)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'              } '						-isItem $True	-indentlevel ($indentDataStart + 4) ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'  } '										-isItem $True	-indentlevel ($indentDataStart + 2)  ))															
					}

					#------- local Accounts
					if ($_localAccountArr)
					{
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'- {'										-isItem $True	-indentlevel ($indentDataStart + 2)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'settingType : LocalAccounts, '			-isItem $True	-indentlevel ($indentDataStart + 4)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'args        : { '						-isItem $True	-indentlevel ($indentDataStart + 4)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			'localAccounts : ['					-isItem $True	-indentlevel ($indentDataStart + 5)  ))									
						
						foreach ($s in $_localAccountArr)
						{
							$privArr 				= @()
							$_uname 				= "userName                 : {0},"		-f $s.username
							$_display				= "displayName              : {0},"		-f $s.displayName
							$_password 				= "password                 : '{0}',"	-f $s.userPassword
							if ($s.userPrivileges)
							{
								$privArr 			= $s.userPrivileges.Split($SepChar)
							}
			
							[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 				'{'								-isItem $True	-indentlevel ($indentDataStart + 6)  ))
							[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 					$_uname						-isItem $True	-indentlevel ($indentDataStart + 8)  ))
							[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 					$_display					-isItem $True	-indentlevel ($indentDataStart + 8)  ))
							[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 					$_password					-isItem $True	-indentlevel ($indentDataStart + 8)  ))
							if ($Null -ne $privArr)
							{
								foreach ($_priv in $privArr)
								{
									$_userPriv 			= $YMLiLOPrivilegeParamEnum.item($_priv) 
									[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			$_userPriv					-isItem $True	-indentlevel ($indentDataStart + 8)  ))
								}
								$YMLscriptCode[-1] = $YMLscriptCode[-1].TrimEnd() -replace ".$"
							}
							else # No user privilege defined 
							{
								$YMLscriptCode[-1] = $YMLscriptCode[-1].TrimEnd() -replace ".$"				# Remove COMMA
							}
							[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 				'},'							-isItem $True	-indentlevel ($indentDataStart + 6)  ))
						}
						$YMLscriptCode[-1] 			= $YMLscriptCode[-1].TrimEnd() -replace ".$"
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			'                ]'					-isItem $True	-indentlevel ($indentDataStart + 5)  ))									
						
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'               } '						-isItem $True	-indentlevel ($indentDataStart + 4)  ))

						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'  }'										-isItem $True	-indentlevel ($indentDataStart + 2)  ))

					}

					#------- Directory Group
					if ($_dirGroupsArr)
					{
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'- {'										-isItem $True	-indentlevel ($indentDataStart + 2)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'settingType : DirectoryGroups , '		-isItem $True	-indentlevel ($indentDataStart + 4)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'args        : { '						-isItem $True	-indentlevel ($indentDataStart + 4)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			'directoryGroupAccounts :   ['		-isItem $True	-indentlevel ($indentDataStart + 5)  ))									
						
						foreach ($s in $_dirGroupsArr)
						{
							$privArr 				= @()
							$_groupDN  				= "groupDN                  : '{0}',"	-f $s.groupDN
							$_groupSID 				= "groupSID                 : {0},"		-f $s.groupSID


							[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 				'{'								-isItem $True	-indentlevel ($indentDataStart + 6)  ))
							[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 					$_groupDN					-isItem $True	-indentlevel ($indentDataStart + 8)  ))
							[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 					$_groupSID					-isItem $True	-indentlevel ($indentDataStart + 8)  ))

							if ($s.groupPrivileges)
							{
								$privArr 			= $s.groupPrivileges.Split($SepChar)
								foreach ($_priv in $privArr)
								{
									$_userPriv 			= $YMLiLOPrivilegeParamEnum.item($_priv) 
									[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			$_userPriv					-isItem $True	-indentlevel ($indentDataStart + 8)  ))
								}
								$YMLscriptCode[-1] = $YMLscriptCode[-1].TrimEnd() -replace ".$"
							}
							else # No user privilege defined 
							{
								$YMLscriptCode[-1] = $YMLscriptCode[-1].TrimEnd() -replace ".$"				# Remove COMMA
							}
							[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 				'},'							-isItem $True	-indentlevel ($indentDataStart + 6)  ))
						}
						$YMLscriptCode[-1] 			= $YMLscriptCode[-1].TrimEnd() -replace ".$"
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			'                           ]'		-isItem $True	-indentlevel ($indentDataStart + 5)  ))									
						
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'              } '						-isItem $True	-indentlevel ($indentDataStart + 4)  ))

						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'  }'										-isItem $True	-indentlevel ($indentDataStart + 2)  ))
							
					}

					#------- KeyManager
					foreach ($s in $_keyManagerArr)
					{
						$_primary 				= "primaryServerAddress:   {0} ,"						-f $s.primaryServerAddress
						$_primaryPort 			= "primaryServerPort:      {0} ,"						-f $s.primaryServerPort 
						if ($s.secondaryServerAddress)
						{
							$_secondary 			= "secondaryServerAddress: {0} ,"					-f $s.secondaryServerAddress
							$_secondaryPort 		= "secondaryServerPort:    {0} ,"					-f $s.secondaryServerPort 
						}
						$_groupName 			= "groupName:              {0} ,"						-f $s.groupName
						$_certificateName 		= "certificateName:        {0} ,"						-f $s.certificateName
						$_loginName 			= "loginName:              {0} ,"						-f $s.loginName
						$_password 				= "password:               '{0}'"						-f $s.keyManagerpassword 

						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'- {'										-isItem $True	-indentlevel ($indentDataStart + 2)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'settingType : KeyManager, '			-isItem $True	-indentlevel ($indentDataStart + 4) ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'args        : { '						-isItem $True	-indentlevel ($indentDataStart + 4) ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			$_primary 							-isItem $True	-indentlevel ($indentDataStart + 5)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			$_primaryPort						-isItem $True	-indentlevel ($indentDataStart + 5)  ))
						if ($s.secondaryServerAddress)
						{
							[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		$_secondary 						-isItem $True	-indentlevel ($indentDataStart + 5)  ))
							[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		$_secondaryPort 					-isItem $True	-indentlevel ($indentDataStart + 5)  ))
						}
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			$_groupName 						-isItem $True	-indentlevel ($indentDataStart + 5)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			$_certificateName					-isItem $True	-indentlevel ($indentDataStart + 5)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			$_loginName 						-isItem $True	-indentlevel ($indentDataStart + 5)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			$_password 							-isItem $True	-indentlevel ($indentDataStart + 5)  ))
						
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'              } '						-isItem $True	-indentlevel ($indentDataStart + 4) ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'  } '										-isItem $True	-indentlevel ($indentDataStart + 2)  ))															
					}

					#------- Directory
					foreach ($s in $_directoryArr)
					{
						$_authentication 		= "directoryAuthentication:    {0}  ,"						-f $YMLiLOdirAuthParamEnum.item($s.directoryAuthentication)
						$_ldap 					= "directoryGenericLDAP:       {0}  ,"						-f $s.directoryGenericLDAP
						$_iloDN 				= "iloObjectDistinguishedName: '{0}',"					    -f $s.iloObjectDistinguishedName
						$_password 				= "password:                   '{0}',"						-f $s.directoryPassword
						$_server 				= "directoryServerAddress:     {0}  ,"						-f $s.directoryServerAddress 
						$_port 					= "directoryServerPort:        {0}  ,"						-f $s.directoryServerPort


						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'- {'										-isItem $True	-indentlevel ($indentDataStart + 2)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'settingType : Directory, '				-isItem $True	-indentlevel ($indentDataStart + 4) ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'args        : { '						-isItem $True	-indentlevel ($indentDataStart + 4) ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			$_authentication 					-isItem $True	-indentlevel ($indentDataStart + 5)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			$_ldap								-isItem $True	-indentlevel ($indentDataStart + 5)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			$_iloDN 							-isItem $True	-indentlevel ($indentDataStart + 5)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			$_password 							-isItem $True	-indentlevel ($indentDataStart + 5)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			$_server 							-isItem $True	-indentlevel ($indentDataStart + 5)  ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			$_port								-isItem $True	-indentlevel ($indentDataStart + 5)  ))
						
						if ($s.directoryUserContext)
						{
							$_userContextArr 	= $s.directoryUserContext.Split($SepChar)
							[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'directoryUserContext: ['  			-isItem $True	-indentlevel ($indentDataStart + 5)  ))

							foreach ($_u in $_userContextArr)
							{
								[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			("'{0}', " -f $_u) 			-isItem $True	-indentlevel ($indentDataStart + 6)  ))
						
							}

							$YMLscriptCode[-1]	= $YMLscriptCode[-1].TrimEnd() -replace ".$"
							[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'                      ]'  			-isItem $True	-indentlevel ($indentDataStart + 5)  ))
							
						}
						
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		'              } '						-isItem $True	-indentlevel ($indentDataStart + 4) ))
						[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'  } '										-isItem $True	-indentlevel ($indentDataStart + 2)  ))															
					}

				}

			}

			# ############### SN-MAC-ISCSI
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	macType					-value $macType				-indentlevel $indentDataStart ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	wwnType					-value $wwnType				-indentlevel $indentDataStart ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	serialNumberType		-value $snType				-indentlevel $indentDataStart ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	iscsiInitiatorNameType	-value $iscsiType			-indentlevel $indentDataStart ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	hideUnusedFlexNics		-value $hideFlexNics		-indentlevel $indentDataStart ))

		}

		# --- Power On server after profile is created
		if (($isSp) -and ($null -ne $serverHardware))
		{
			newLine	-code $YMLscriptCode
			$title 			= ' Power on server {0}' 	-f $serverHardware
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix name  				-value $title	-isVar $True 				-indentlevel 1 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix oneview_server_hardware 										-indentlevel 1 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	config				-value "'{{config}}'"					-indentlevel 2 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	state				-value "power_state_set"				-indentlevel 2 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 	'data'														-indentlevel 2 ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		name				-value "'$serverHardware'"			-indentlevel $indentDataStart )) 
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 		powerStateData 											-indentlevel $indentDataStart ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			powerState			-value "'On'" 					-indentlevel ($indentDataStart + 1) ))
			[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix 			powerControl		-value "'MomentaryPress'" 		-indentlevel ($indentDataStart + 1) ))
		}

	}

	#newLine	-code $YMLscriptCode
	#[void]$YMLscriptCode.Add((Generate-YMLCustomVarCode -prefix delegate_to				-value localhost									-indentlevel 2 ))

	 # ---------- Generate script to file
	 YMLwriteToFile -code $YMLscriptCode -file $YMLfiles

}


# -------------------------------------------------------------------------------------------------------------
#
#       Main Entry
#
# -------------------------------------------------------------------------------------------------------------



# ---------------- Define Excel files
#
$startRow 				= 15

$allPSscriptCode 			= [System.Collections.ArrayList]::new()

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
			$dir    	= "$currentDir\$dir"
			$ansibledir = "ansible_playbook"
			if (-not (test-path -path $dir) )
			{
                write-host -ForegroundColor Cyan "--------- Creating folder $dir"
                write-host -ForegroundColor Cyan $CR
				md $dir
				md "$dir\$ansibledir"
			}
		}

		$allScriptFile			= "$currentDir\allScripts.ps1"



		$sheetNames 	= (Get-ExcelSheetInfo -Path $workBook).Name
		$composer 		= 'OVdestination'
		$sequence 		= 1


		#----------------------------------------------
		#              OV Resources
		#----------------------------------------------
		[void]$allPSscriptCode.Add($CR)
		[void]$allPSscriptCode.Add('#-----------------------------------------------')
		[void]$allPSscriptCode.Add('#              OV Resources     				')
		[void]$allPSscriptCode.Add('#-----------------------------------------------')
		[void]$allPSscriptCode.Add($CR)

		# ================ Appliance folder
		$subdir         = "$currentdir\Appliance"
		$subAnsibledir	= "$subdir\$ansibledir"

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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
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
		$subAnsibledir	= "$subdir\$ansibledir"

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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++

				$_sheet		= $sheet +"_ipV4"
				$YMLfiles 	= "$subAnsibledir\$_sheet.yml"

				Import-YMLaddressPool_ipv4 -sheetName $sheetName -workBook $workbook -YMLfiles $YMLFiles 
				write-host -ForegroundColor Cyan "Ansible playbook is created ---> $YMLfiles "
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
		$subAnsibledir	= "$subdir\$ansibledir"

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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++

				$YMLfiles 	= "$subAnsibledir\$sheet.yml"
				Import-YMLethernetNetwork -sheetName $sheetName -workBook $workbook -YMLfiles $YMLfiles
				write-host -ForegroundColor Cyan "Ansible playbook is created ---> $YMLfiles "
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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++

				$YMLfiles 	= "$subAnsibledir\$sheet.yml"
				Import-YMLfcNetwork -sheetName $sheetName -workBook $workbook -YMLfiles $YMLfiles
				write-host -ForegroundColor Cyan "Ansible playbook is created ---> $YMLfiles "
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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++

				$YMLfiles 	= "$subAnsibledir\$sheet.yml"
				Import-YMLnetworkSet -sheetName $sheetName -workBook $workbook -YMLfiles $YMLfiles
				write-host -ForegroundColor Cyan "Ansible playbook is created ---> $YMLfiles "
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
				$sheetArray 	= 'logicalInterconnectGroup|uplinkSet|snmpConfiguration|snmpV3User|snmpTrap'.Split($SepChar)
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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++

				$YMLfiles 	= "$subAnsibledir\$sheet.yml"
				Import-YMLLogicalInterconnectGroup -sheetName $sheetName -workBook $workbook -YMLfiles $YMLfiles
				write-host -ForegroundColor Cyan "Ansible playbook is created ---> $YMLfiles "

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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++

				$YMLfiles 	= "$subAnsibledir\$sheet.yml"
				Import-YMLUplinkSet -sheetName $sheetName -workBook $workbook -YMLfiles $YMLfiles
				write-host -ForegroundColor Cyan "Ansible playbook is created ---> $YMLfiles "
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

		# ---- Import LIG snmp
		$sheet 			= 'snmpConfiguration' 
		$resource 		= 'logical interconnect group - snmp'
		
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				$sheetArray 	= 'snmpConfiguration|snmpV3User|snmpTrap'.Split($SepChar)
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
				$sequence++
				$YMLfiles 	= "$subAnsibledir\lig-$sheet.yml"
				Import-YMLligSNMP -sheetName $sheetName -workBook $workbook -YMLfiles $YMLfiles
				write-host -ForegroundColor Cyan "Ansible playbook is created ---> $YMLfiles "

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
		$subAnsibledir	= "$subdir\$ansibledir"


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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
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
		$subAnsibledir	= "$subdir\$ansibledir"


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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++

				$YMLfiles       = "$subAnsibledir\$sheet.yml"
				Import-YMLEnclosureGroup -sheetName $sheetName -workBook $workbook -YMLfiles $YMLFiles 
				write-host -ForegroundColor Cyan "Ansible playbook is created ---> $YMLfiles "
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
				# Will split 1 PS1 per LE
				$ListPS1Files = Import-LogicalEnclosure -sheetName $sheetName -workBook $workbook -subdir $subdir

				foreach ($ps1Files in $ListPS1Files)
				{
					write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
					add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				}
				$sequence++

				
				# Will split 1 YML per LE
				$ListYMLfiles = Import-YMLlogicalEnclosure -sheetName $sheetName -workBook $workbook -subdir $subAnsibledir
				foreach ($YMLfiles in $ListYMLfiles)
				{
					write-host -ForegroundColor Cyan "Script is created           ---> $YMLfiles "
				}
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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
				add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				$sequence++

				$YMLfiles       = "$subAnsibledir\$sheet.yml"
				Import-YMLProfileorTemplate -sheetName $sheetName -workBook $workbook -YMLFiles $YMLfiles -isSpt $True -subdir $subdir
				write-host -ForegroundColor Cyan "Ansible playbook is created ---> $YMLfiles "
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
					write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
					add_to_allScripts -ps1Files $ps1Files -text "# ------ Step $sequence - Create $resource script"
				}
				$sequence++

				$YMLfiles       = "$subAnsibledir\$sheet.yml"
				Import-YMLProfileorTemplate -sheetName $sheetName -workBook $workbook -YMLFiles $YMLfiles -isSpt $False -subdir $subdir
				write-host -ForegroundColor Cyan "Ansible playbook is created ---> $YMLfiles "
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
		$subAnsibledir	= "$subdir\$ansibledir"

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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
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
		


		

		[void]$allPSscriptCode.Add($CR)
		[void]$allPSscriptCode.Add('#----------------------------------------------')
		[void]$allPSscriptCode.Add('#              OV Settings					 ')
		[void]$allPSscriptCode.Add('#----------------------------------------------')
		[void]$allPSscriptCode.Add($CR)

		# ================ Appliance folder
		$subdir         = "$currentdir\Appliance"
		$subAnsibledir	= "$subdir\$ansibledir"

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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
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
		$subAnsibledir	= "$subdir\$ansibledir"


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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
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
		$subAnsibledir	= "$subdir\$ansibledir"


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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
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
				write-host -ForegroundColor Cyan "Script is created           ---> $ps1Files "
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

		# ---- Import Appliance snmp
		$sheet 			= 'snmpConfiguration' 
		$resource 		= 'Appliance snmp'
		
		if ($sheetNames -contains $sheet)
		{
			$data 		= import-Excel -workSheetName $sheet -path $workbook -StartRow ($startRow -1) -noHeader
			if ($data.Count -gt 0)
			{
				$sheetArray 	= 'snmpConfiguration|snmpV3User|snmpTrap'.Split($SepChar)
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
				$sequence++
				$YMLfiles 	= "$subAnsibledir\$sheet.yml"
				Import-YMLSNMP -sheetName $sheetName -workBook $workbook -YMLfiles $YMLfiles
				write-host -ForegroundColor Cyan "Ansible playbook is created ---> $YMLfiles "
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

		[void]$allPSscriptCode.Add($CR)
		[void]$allPSscriptCode.Add('#----------------------------------------------')
        [void]$allPSscriptCode.Add('#              TBD -OV Appliance configuration		 ')
		[void]$allPSscriptCode.Add('#----------------------------------------------')
		[void]$allPSscriptCode.Add($CR)




		# ----- Generate all Scripts file
		write-host -ForegroundColor Cyan "`n`n--------- All-in-one script"
		write-host -ForegroundColor CYan "`n$allScriptFile contains all individual scripts that can be run to configure your new environment.`n`n"
		writeToFile -code $allPSscriptCode -file $allScriptFile
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






