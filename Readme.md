<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128610451/11.2.5%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/E1680)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [Form1.cs](./CS/MailMerge/Form1.cs) (VB: [Form1.vb](./VB/MailMerge/Form1.vb))
* [SampleData.cs](./CS/MailMerge/SampleData.cs) (VB: [SampleData.vb](./VB/MailMerge/SampleData.vb))
<!-- default file list end -->
# How to implement mail merge in a RichEditControl


<p>This example illustrates an older approach to implement mail merge in the <strong>RichEditControl</strong>.<br />
To learn about a newer and more convenient approach to <a href="https://documentation.devexpress.com/#WindowsForms/CustomDocument16044"><u>master-detail mail merge</u></a>, refer to <a href="https://www.devexpress.com/Support/Center/CodeCentral/ViewExample.aspx?exampleId=E5078"><u>How to automatically create mail-merge documents using the Snap Document Server</u></a>.</p><p><br />
In this example, the <strong>ArrayList</strong> that is generated at runtime is used as a data source that supplies mail merge data to the document. The tab control on the form contains a <a href="https://documentation.devexpress.com/#WindowsForms/CustomDocument9551"><u>Ribbon UI</u></a> and two <strong>RichEditControl</strong> instances (one of them is used to construct a document template, and the other displays the mail merge result).</p>

<br/>


