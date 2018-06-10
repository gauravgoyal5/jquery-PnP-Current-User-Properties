import * as pnp from 'sp-pnp-js';

import { Version, Environment, EnvironmentType, DisplayMode } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';

import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GetCurrentUserPropertiesWithPnPWebPart.module.scss';
import * as strings from 'GetCurrentUserPropertiesWithPnPWebPartStrings';



export interface IGetCurrentUserPropertiesWithPnPWebPartProps {
  description: string;
}

export default class GetCurrentUserPropertiesWithPnPWebPart extends BaseClientSideWebPart<IGetCurrentUserPropertiesWithPnPWebPartProps> {

  public strLog: string = '';
  private strImagePath: string='';
  public onInit(): Promise<void> {
    let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
    SPComponentLoader.loadCss(cssURL);
    return Promise.resolve<void>();
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.getCurrentUserPropertiesWithPnP}">
        
        <div class="container">
        <div class="row">
            <div class="col-md-6 col-sm-8">
                <div class="${ styles.pricingTable}">
                    <div class="${ styles["pricingTable-header"]}">
                        <h3 class="title">My Profile Information</h3>
                    </div>
                    <div class="${ styles["pricing-content"]}">
                        <div class="${ styles["price-value"]}">
                          <img id="myImage" >
                            
                        </div>
                        <ul id="spUserProfileProperties">
                          
                        </ul>
                    </div>
                </div>
            </div>
            </div>
      </div>`;
      this.getCurrentUserProfileProperties();
  }

  private getCurrentUserProfileProperties(): string {
    let htmlUserProperties: string = "";
    this.strLog += "<br>1";
    if (Environment.type == EnvironmentType.Local) {
      this.strLog += "<br>2";
      htmlUserProperties = "<h1>You are workin in local workbench</h1>";
      return htmlUserProperties;
    }

    this.strLog += "<br>3";
    pnp.sp.profiles.myProperties.get().then((result) => {
      this.strLog += "<br>4";
      var userProperties = result.UserProfileProperties;
      this.strLog += "<br>5";

      userProperties.forEach((retVal) => {
        this.strLog += "<br>" + retVal.Key;
        htmlUserProperties += "<li><strong>" + retVal.Key + "</strong> - " + retVal.Value + "</li>";
        if(retVal.Key=="PictureURL")
        {
          (document.getElementById("myImage") as HTMLImageElement).src=retVal.Value;
        }
      });
      this.strLog += "<br>6";

      document.getElementById("spUserProfileProperties").innerHTML = htmlUserProperties;
    }).catch((error) => {
      //console.error("Error: " + error);
      //alert(error);
      this.strLog += "<br>7";
      this.strLog += "<br>Error " + error;
      return (htmlUserProperties = error);
    });
    this.strLog += "<br>8";
    return htmlUserProperties;

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
