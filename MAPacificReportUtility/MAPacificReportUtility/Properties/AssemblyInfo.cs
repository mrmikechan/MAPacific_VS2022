using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("MAPacificReportUtility")]
[assembly: AssemblyDescription("Custom Macro to process reports with OutsideView. Current version addresses: Migrate away from Exchange Webservices to Graph API to handle Azure AD Authentication and email drafts")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("Crystal Point, inc")]
[assembly: AssemblyProduct("MAPacificReportUtility")]
[assembly: AssemblyCopyright("Copyright © Crystal Point, inc 2022")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("3c854d63-2a68-4063-b1ab-dfeb651ce66c")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Build and Revision Numbers 
// NOTE:  The version number represents the phases of development in the MAPacific Project docx. For instance Phase 7
// in the document is 1.7 in the Assembly version!
// by using the '*' as shown below:
// [assembly: AssemblyVersion("1.0.*")]
//updated version to 1.10.9 -- 3/27/22 -- Version fixes bug in which office365/outlook365 requires that you set security protocol to TSL12 before using the Web Service.
//Updated version to 1.11.0 -- 7/4/22 -- Version migrates away from Exchange Web Services to Graph API for Authentication and email drafts
[assembly: AssemblyVersion("1.11.0.0")]
[assembly: AssemblyFileVersion("1.11.0.0")]
