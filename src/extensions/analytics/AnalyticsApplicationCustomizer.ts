import { spfi, SPFx } from "@pnp/sp";
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import * as strings from 'AnalyticsApplicationCustomizerStrings';

const LOG_SOURCE: string = 'AnalyticsApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAnalyticsApplicationCustomizerProperties {
  // This is an example; replace with your own property
  userEmail: string;
  externalUser: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AnalyticsApplicationCustomizer
  extends BaseApplicationCustomizer<IAnalyticsApplicationCustomizerProperties> {

    private previousUrl: string = '';

    public async onInit(): Promise<void> {
      Log.info(LOG_SOURCE, 'Initialized AnalyticsApplicationCustomizer');
      console.log("Initialized Analytics Extension");

      // Initialize PnP JS
      const sp = spfi().using(SPFx(this.context));

      // Track the initial page load
      //await this.handlePageTracking();

      const user = await this.context.msGraphClientFactory.getClient("3").then((client: MSGraphClientV3): Promise<any> => {
        return client.api('/me').get();
      });

      const userName = user.displayName;
      this.properties.userEmail = user.mail;
      const useremail = user.mail;
      const companyname = this._getCompanyFromEmail();
      const siteID = this.context.pageContext.site.id.toString();
      const pageID = this.context.pageContext.legacyPageContext["pageItemId"];
      const currentDatetime = new Date().toISOString(); // ISO 8601 format
      const usercompany = this._getCompanyFromEmail();
      const devicecontext = this._getDeviceContext();

      const logData = {
        userName,
        useremail,
        companyname,
        siteID,
        pageID,
        currentDatetime,
        usercompany,
        devicecontext
      };

      console.log('Page Tracking Data:', logData);
      //
      // Send the page tracking data to the provided API
      await this.sendPageTrackingData(logData);

      // Monitor URL changes
      this.monitorUrlChanges();

      return Promise.resolve();
    }

    private async monitorUrlChanges(): Promise<void> {
      // Track URL changes on page load
      window.addEventListener('load', async () => {
        console.log("new page is loaded");
        await this.handlePageTracking();
      });

     // Handle URL change on page navigation
      /*window.addEventListener('popstate', async () => {
        console.log("popstate activated");
        await this.handlePageTracking();
      });*/

      // Handle URL change due to pushState or replaceState
      
      const originalPushState = history.pushState;
      history.pushState = (...args) => {
        console.log("pushstate activated");
        originalPushState.apply(history, args);
        this.handlePageTracking();
      };

      const originalReplaceState = history.replaceState;
      history.replaceState = (...args) => {
        console.log("replacestate activated");
        originalReplaceState.apply(history, args);
        this.handlePageTracking();
      };
}

    private async handlePageTracking(): Promise<void> {
      const currentUrl = window.location.href;
      if (this.isUrlChanged(currentUrl)) {
        // Get the current user and other tracking details
        const user = await this.context.msGraphClientFactory.getClient("3").then((client: MSGraphClientV3): Promise<any> => {
          return client.api('/me').get();
        });

        const userName = user.displayName;
        this.properties.userEmail = user.mail;
        const useremail = user.mail;
        const companyname = this._getCompanyFromEmail();
        const siteID = this.context.pageContext.site.id.toString();
        const pageID = this.context.pageContext.legacyPageContext["pageItemId"];
        const currentDatetime = new Date().toISOString(); // ISO 8601 format
        const usercompany = this._getCompanyFromEmail();
        const devicecontext = this._getDeviceContext();

        const logData = {
          userName,
          useremail,
          companyname,
          siteID,
          pageID,
          currentDatetime,
          usercompany,
          devicecontext
        };

        console.log('Page Tracking Data:', logData);

        // Send the page tracking data to the provided API
        await this.sendPageTrackingData(logData);

        // Update the previous URL
        this.previousUrl = window.location.href.split("https://rpgnet.sharepoint.com/sites/OneRPG")[1] || "";
      }
    }

    private isUrlChanged(currentUrl: string): boolean {
      const baseUrl = "https://rpgnet.sharepoint.com/sites/OneRPG";
      const urlPart = currentUrl.split(baseUrl)[1] || "";
      return this.previousUrl !== urlPart;
    }

    private async sendPageTrackingData(pageTrackingData: any): Promise<void> {
      const apiEndpoint = 'https://azfunctionrpguserusagereport.azurewebsites.net/api/UserUsageReportsFunction?code=xaM5TIC0fO0T7rIzJezNJvO3EyTVmiSbyVcbgtpH263CAzFuGIYgRQ=='; 

      try {
        const response = await fetch(apiEndpoint, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Accept': 'application/json'
          },
          body: JSON.stringify({
            username: pageTrackingData.userName,
            useremail: pageTrackingData.useremail,
            companyname: pageTrackingData.companyname,
            siteID: pageTrackingData.siteID,
            pageID: pageTrackingData.pageID,
            currentDatetime: pageTrackingData.currentDatetime,
            usercompany: pageTrackingData.usercompany,
            devicecontext: pageTrackingData.devicecontext
          })
        });

        if (response.ok) {
          console.log('Page tracking data sent successfully');
        } else {
          console.error('Failed to send page tracking data:', response.statusText);
        }
      } catch (error) {
        console.error('Error sending page tracking data:', error);
      }
    }

    private _getCompanyFromEmail(): string {
      let userGroup: string = "";

      if(this.properties.userEmail.includes("@aciesinnovations.com")) {
        userGroup = "acies innovations";
      } else if(this.properties.userEmail.includes("_zensar.com") || this.properties.userEmail.includes("@zensar.")) {
        userGroup = "zensar";
      } else if(this.properties.userEmail.includes("@ceat.com")) {
        userGroup = "ceat";
      } else if(this.properties.userEmail.includes("_harrisonsmalayalam.com") || this.properties.userEmail.includes("@harrisonsmalayalam.com")) {
        userGroup = "harrison";
      } else if(this.properties.userEmail.includes("@kecrpg.com")) {
        userGroup = "kec";
      } else if(this.properties.userEmail.includes("@raychemrpg.com")) {
        userGroup = "raychem";
      } else if(this.properties.userEmail.includes("@rpgls.com")) {
        userGroup = "rpgls";
      } else if(this.properties.userEmail.includes("@rpg.com") || this.properties.userEmail.includes("@rpg.in")) {
        userGroup = "rpg";
      }
      return userGroup;
    }  

    private _getDeviceContext(): string {
      const userAgent = navigator.userAgent;
      // logic to determine device context
      if (/Mobi|Android/i.test(userAgent)) {
        return 'Mobile';
      } else if (/iPad|Tablet/i.test(userAgent)) {
        return 'Tablet';
      } else {
        return 'Desktop';
      }
    }
}
