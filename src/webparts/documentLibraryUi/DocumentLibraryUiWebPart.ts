import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import styles from './DocumentLibraryUiWebPart.module.scss';
import * as strings from 'DocumentLibraryUiWebPartStrings';
import MockHttpClient from './MockHttpClient'; 
import { SPHttpClient , SPHttpClientResponse } from '@microsoft/sp-http';
import {  
  Environment,  
  EnvironmentType  
} from '@microsoft/sp-core-library';

export interface IDocumentLibraryUiWebPartProps {
  description: string;
}
export interface ISPLists {  
  value: ISPList[];  
}  
export interface ISPList {  
  FileLeafRef: string;  
  // Name: string; 
  DocTitle: string;  
  DescriptionOfDoc: string;  
  Department:string;
}    

export default class DocumentLibraryUiWebPart extends BaseClientSideWebPart<IDocumentLibraryUiWebPartProps> {

  private _getMockListData(): Promise<ISPLists> {  
    return MockHttpClient.get(this.context.pageContext.web.absoluteUrl).then(() => {  
        const listData: ISPLists = {  
            value:  
            [  
                // { Name: 'E123', EmployeeName: 'John', Experience: 'SharePoint',Location: 'India' },  
                //  { Name: 'E567', EmployeeName: 'Martin', Experience: '.NET',Location: 'Qatar' },  
                // { Name: 'E367', EmployeeName: 'Luke', Experience: 'JAVA',Location: 'UK' }  
            ]  
            };  
        return listData;  
    }) as Promise<ISPLists>;  
}   

private _getListData(): Promise<ISPLists> {  
  
  return this.context.spHttpClient.get(`https://epoms.sharepoint.com/_api/web/lists/getbytitle('DocumentedInformation')/Items?$select=FileLeafRef,DocTitle,Department,DescriptionOfDoc`, SPHttpClient.configurations.v1)  
      .then((response: SPHttpClientResponse) => {   
        console.log('response',response)
        // debugger;  
        return response.json();  
      });  
  }   

  private _renderListAsync(): void {  
      
    if (Environment.type === EnvironmentType.Local) {  
      this._getMockListData().then((response) => {  
        this._renderList(response.value);  
      });  
    }  
     else {  
       this._getListData()  
      .then((response) => {  
        this._renderList(response.value);  
      });  
   }  
} 

private _renderList(items: ISPList[]): void {  
  if(items.length>0)
  {
    
    let no=0;
  let html: string;
  html+= `<table style="width:100%!important">`;
  html+=` <thead><tr class="${styles.th}"  valign="top"><th>No</th><th  width="25%">Title</th><th>Department</th><th width="40%">Description</th><th>Attachments</th></tr></thead>`;  
  items.forEach((item: ISPList) => { 
    no++;
      var url = item.FileLeafRef;
      var dept = item.Department;
      var title = item.DocTitle;
      var desc = item.DescriptionOfDoc;
      // var ver = item.DocumentVersion;
      // var docid=item.DocumentID;
      var jsonString;
      jsonString=url.split('/');
    
      //var arrlen=jsonString.length;
    
      html+=`<tbody><tr class="${styles.th1}"  valign="top">`;    
       
         html+=` <td  data-label="No"><div>${no}</div></td>  `;
      
        if(title==null)
        {
         html+=` <td data-label="Title"><div><b style="color:#ececec;">-</b></div></td>  `;
        }
        else
        {
         html+=` <td data-label="Title"><div>${title}</div></td>  `;
        }

        if(dept==null)
        {
         html+=` <td  data-label="Department"><div><b style="color:#ececec;">-</b></div></td>  `;
        }
        else
        {
         html+=` <td  data-label="Department"><div>${dept}</div></td>  `;
        }

         if(desc==null)
         {
          html+=` <td data-label="Description"><div><b style="color:#ececec;">-</b></div></td>  `;
         }
         else
         {
          html+=` <td data-label="Description" ><div>${desc}</div></td>  `;
         }
         
         if(url==null)
        {
          html+=` <td data-label="Attachment"><div><b style="color:#ececec;">-</b></div></td>  `;
        }
        else
        {  
          var imagetype=url.lastIndexOf('.');
          var image=url.substring(imagetype);
          var imgtype;
          var pdfex=".pdf";
          var wordex=".docx";
          var docex=".doc";
          var excelex=".xlsx";
          imgtype= image.toLowerCase( )
          if(imgtype.toString()==pdfex.toString())
          {
            
            html+=`<td data-label="Attachment"><div><a  onclick='window.open("https://epoms.sharepoint.com/Documented%20Information/${jsonString}","_blank")'><img src="https://epoms.sharepoint.com/SiteAssets/documentlibraryicon/pdf.png" width="50px" height="50px"></a></div></td>   `;
          }
          else if(imgtype.toString()==wordex.toString())
          {
            html+=`<td data-label="Attachment"><div><a  onclick='window.open("https://epoms.sharepoint.com/Documented%20Information/${jsonString}","_blank")'><img src="https://epoms.sharepoint.com/SiteAssets/documentlibraryicon/doc.png" width="50px" height="50px"></a></div></td>   `;
          }
          else if(imgtype.toString()==docex.toString())
          {
            html+=`<td data-label="Attachment"><div><a  onclick='window.open("https://epoms.sharepoint.com/Documented%20Information/${jsonString}","_blank")'><img src="https://epoms.sharepoint.com/SiteAssets/documentlibraryicon/doc.png" width="50px" height="50px"></a></div></td>   `;
          }
          else if(imgtype.toString()==excelex.toString())
          {
            html+=`<td data-label="Attachment"><div><a  onclick='window.open("https://epoms.sharepoint.com/Documented%20Information/${jsonString}","_blank")'><img src="https://epoms.sharepoint.com/SiteAssets/documentlibraryicon/xlsx.png" width="50px" height="50px"></a></div></td>   `;
          }
          else
          {
            console.log(jsonString)
            html+=`<td data-label="Attachment"><div><a  onclick='window.open("https://epoms.sharepoint.com/Documented%20Information/${jsonString}","_blank")'><img src="https://epoms.sharepoint.com/SiteAssets/documentlibraryicon/folder.jpg" width="50px" height="50px"></a></div></td>   `;
          }
          
         
        }
        
       
        
        html+=`</tr> </tbody> `;  
        
          }); 
        html += `</table>`; 

  const listContainer: Element = this.domElement.querySelector('#spListContainer');  
  listContainer.innerHTML = html;  
        }
      
}  

public render(): void {  
  this.domElement.innerHTML = `  
  <div class="${styles.documentLibraryUi}">  
  <div class="${styles.container}">  
  <div class="${styles.row}">  
  
<div id="spListContainer" />  
  
</div>
</div>
</div>
`;  
this._renderListAsync();  
}  
  

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
