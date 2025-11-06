import { spfi, SPFx } from "@pnp/sp";
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

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
    this.properties.userEmail = user.mail.toLowerCase();
    const useremail = user.mail;
    const companyname = this._getCompanyFromEmail();
    const siteID = this.context.pageContext.site.id.toString();
    const pageID = this.context.pageContext.legacyPageContext["pageItemId"];
    const url = this.context.pageContext.site.serverRequestPath;
    const pageName = url.substring(url.lastIndexOf('/') + 1, url.lastIndexOf('.aspx'));
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
      devicecontext,
      pageName
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
    console.log("🚀 monitorUrlChanges() called");

    // Handle URL change due to pushState or replaceState
    const originalPushState = history.pushState;
    history.pushState = (...args) => {
      console.log("🔄 PUSHSTATE ACTIVATED");
      originalPushState.apply(history, args);
      this.handlePageTracking();
    };

    const originalReplaceState = history.replaceState;
    history.replaceState = (...args) => {
      console.log("🔁 REPLACESTATE ACTIVATED");
      originalReplaceState.apply(history, args);
      this.handlePageTracking();
    };

    // Trigger initial tracking only once
    console.log("🎯 TRIGGERING INITIAL TRACKING");
    await this.handlePageTracking();
  }

  private async logToMailerClicks(currentUrl: string): Promise<void> {
    try {
      const urlObject = new URL(currentUrl);
      const urlParams = new URLSearchParams(urlObject.search);
      let issueDate = urlParams.get('issue') || '';

      // 🧠 Format issue date from "September222025" → "September 22, 2025"
      issueDate = this.formatIssueDate(issueDate);

      console.log('🔍 Full URL:', currentUrl);
      console.log('🔍 Extracted & formatted Issue Date:', issueDate);

      // Ensure userEmail is set
      if (!this.properties.userEmail) {
        const user = await this.context.msGraphClientFactory.getClient("3").then((client: MSGraphClientV3): Promise<any> => {
          return client.api('/me').get();
        });
        this.properties.userEmail = user.mail.toLowerCase();
      }

      const sp = spfi().using(SPFx(this.context));

      // Get page details
      const url = this.context.pageContext.site.serverRequestPath;
      const pageName = url.includes('.aspx')
        ? url.substring(url.lastIndexOf('/') + 1, url.lastIndexOf('.aspx'))
        : (this.context.pageContext.legacyPageContext["pageTitle"] || document.title);
      const pageID = this.context.pageContext.legacyPageContext["pageItemId"];
      const pageIDValue = pageID ? parseInt(pageID.toString()) : 0;

      const itemData = {
        Title: this.properties.userEmail,
        Issue: issueDate,
        AccessedDateAndTime: new Date().toISOString(), // Full ISO timestamp
        AccessedPageURL: currentUrl,
        AccessedPageID: pageIDValue,
        AccessedPageName: pageName
      };

      console.log('📝 Data being saved to Analytics_MailerClicks:', itemData);

      const item = await sp.web.lists.getByTitle("Analytics_MailerClicks").items.add(itemData);

      console.log('✅ Saved to Analytics_MailerClicks:', item.data);
    } catch (error) {
      console.error('❌ Error logging to Analytics_MailerClicks:', error);
    }
  }

  private async logToNewJoineeMailerClicks(currentUrl: string): Promise<void> {
    try {
      const urlObject = new URL(currentUrl);
      const urlParams = new URLSearchParams(urlObject.search);
      let issue = urlParams.get('issue') || '';
      const company = urlParams.get('company') || urlParams.get('Company') || '';
      const userProfile = urlParams.get('userprofile') || urlParams.get('AccessedUserProfile') || '';

      // 🧠 Format issue date if same pattern
      issue = this.formatIssueDate(issue);

      console.log('🔍 Full URL:', currentUrl);
      console.log('🔍 Extracted Issue:', issue);
      console.log('🔍 Company:', company);
      console.log('🔍 UserProfile:', userProfile);

      if (!this.properties.userEmail) {
        const user = await this.context.msGraphClientFactory.getClient("3").then((client: MSGraphClientV3): Promise<any> => {
          return client.api('/me').get();
        });
        this.properties.userEmail = user.mail.toLowerCase();
      }

      const sp = spfi().using(SPFx(this.context));

      const itemData = {
        Title: this.properties.userEmail,
        Issue: issue,
        Company: company,
        AccessedUserProfile: userProfile,
        AccessedDateAndTime: new Date().toISOString()
      };

      console.log('📝 Data being saved to Analytics_NewJoineeMailerClicks:', itemData);

      const item = await sp.web.lists.getByTitle("Analytics_NewJoineeMailerClicks").items.add(itemData);

      console.log('✅ Saved to Analytics_NewJoineeMailerClicks:', item.data);
    } catch (error) {
      console.error('❌ Error logging to Analytics_NewJoineeMailerClicks:', error);
    }
  }

  /**
   * Helper: Converts issue string like "September222025" → "September 22, 2025"
   */
  private formatIssueDate(issue: string): string {
    try {
      const match = issue.match(/^([A-Za-z]+)(\d{1,2})(\d{4})$/);
      if (match) {
        const [, month, day, year] = match;
        return `${month} ${day}, ${year}`;
      }
      return issue; // return original if not in expected format
    } catch {
      return issue;
    }
  }


  private async handlePageTracking(): Promise<void> {
    // For testing - hardcode your URL here
    const currentUrl = "https://rpgnet.sharepoint.com/sites/OneRPG/SitePages/Rewards-&-Recognition-Q1--2025-2026.aspx?source=mailer&issue=September222025";
    // const currentUrl = "https://rpgnet.sharepoint.com/sites/OneRPG/SitePages/Rewards-&-Recognition-Q1--2025-2026.aspx?source=organnouncement&issue=September222025&company=RPG&userprofile=profile123";
    // const currentUrl = window.location.href; // Use this for production

    console.log('🔍 Current URL:', currentUrl);

    // Check URL parameters first (before isUrlChanged check)
    if (currentUrl.includes('source=mailer')) {
      console.log('🔵 MAILER LINK DETECTED - Logging to Analytics_MailerClicks');
      await this.logToMailerClicks(currentUrl);
    } else if (currentUrl.includes('source=organnouncement')) {
      console.log('🟢 ORGANNOUNCEMENT LINK DETECTED - Logging to Analytics_NewJoineeMailerClicks');
      await this.logToNewJoineeMailerClicks(currentUrl);
    } else {
      console.log('⚪ REGULAR LINK - No special tracking required');
    }

    // Only proceed with full tracking if URL changed
    if (this.isUrlChanged(currentUrl)) {
      console.log('🔄 URL has changed, performing regular tracking');

      // Get the current user
      const user = await this.context.msGraphClientFactory.getClient("3").then((client: MSGraphClientV3): Promise<any> => {
        return client.api('/me').get();
      });

      const userName = user.displayName;
      this.properties.userEmail = user.mail.toLowerCase();
      const useremail = user.mail;
      const companyname = this._getCompanyFromEmail();
      const siteID = this.context.pageContext.site.id.toString();
      const pageID = this.context.pageContext.legacyPageContext["pageItemId"];
      const url = this.context.pageContext.site.serverRequestPath;
      const pageName = url.includes('.aspx')
        ? url.substring(url.lastIndexOf('/') + 1, url.lastIndexOf('.aspx'))
        : (this.context.pageContext.legacyPageContext["pageTitle"] || document.title);
      const currentDatetime = new Date().toISOString();
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
        devicecontext,
        pageName
      };

      console.log('📊 Page Tracking Data:', logData);
      await this.sendPageTrackingData(logData);

      this.previousUrl = window.location.href.split("https://rpgnet.sharepoint.com/sites/OneRPG")[1] || "";
    } else {
      console.log('⏭️ URL unchanged, skipping regular tracking');
    }
  }

  private isUrlChanged(currentUrl: string): boolean {
    const baseUrl = "https://rpgnet.sharepoint.com/sites/OneRPG";
    const urlPart = currentUrl.split(baseUrl)[1] || "";
    return this.previousUrl !== urlPart;
  }

  private async sendPageTrackingData(pageTrackingData: any): Promise<void> {
    const apiEndpoint = 'https://azfunctionrpguseranalyticsreport.azurewebsites.net/api/UserUsageReportsFunction?code=dCjmkrTo7pu5czZE9UuYCghykJLVZ7EMLEy5RIEWAQy9AzFuy4dP7Q==';

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
          pageName: pageTrackingData.pageName,
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

    if (this.properties.userEmail.includes("@aciesinnovations.com")) {
      userGroup = "acies innovations";
    } else if (this.properties.userEmail.includes("_zensar.com") || this.properties.userEmail.includes("@zensar.")) {
      userGroup = "zensar";
    } else if (this.properties.userEmail.includes("@ceat.com")) {
      userGroup = "ceat";
    } else if (this.properties.userEmail.includes("_harrisonsmalayalam.com") || this.properties.userEmail.includes("@harrisonsmalayalam.com")) {
      userGroup = "harrison";
    } else if (this.properties.userEmail.includes("@kecrpg.com")) {
      userGroup = "kec";
    } else if (this.properties.userEmail.includes("@raychemrpg.com")) {
      userGroup = "raychem";
    } else if (this.properties.userEmail.includes("@rpgls.com")) {
      userGroup = "rpgls";
    } else if (this.properties.userEmail.includes("@rpg.com") || this.properties.userEmail.includes("@rpg.in")) {
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
