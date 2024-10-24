// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { Component, ElementRef, ViewChild } from '@angular/core';
import { IReportEmbedConfiguration, models, Page, Report, service, Embed } from 'powerbi-client';
import { PowerBIReportEmbedComponent } from 'powerbi-client-angular';
import { IHttpPostMessageResponse } from 'http-post-message';
import 'powerbi-report-authoring';

import { reportUrl } from '../public/constants';
import { HttpService } from './services/http.service';

// Handles the embed config response for embedding
export interface ConfigResponse {
  Id: string;
  embedReports :{
    embedUrl: string;
  };
  embedToken: string;
}

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  // Wrapper object to access report properties
  @ViewChild(PowerBIReportEmbedComponent) reportObj!: PowerBIReportEmbedComponent;

  // Track Report embedding status
  isEmbedded = false;

  // Overall status message of embedding
  displayMessage = 'The report is bootstrapped. Click Embed Report button to set the access token.';

  // CSS Class to be passed to the wrapper
  reportClass = 'report-container';

  // Flag which specify the type of embedding
  phasedEmbeddingFlag = false;

  // Pass the basic embed configurations to the wrapper to bootstrap the report on first load
  // Values for properties like embedUrl, accessToken and settings will be set on click of button
  reportConfig: IReportEmbedConfiguration = {
    type: 'report',
    embedUrl: undefined,
    tokenType: models.TokenType.Embed,
    accessToken: undefined,
    settings: undefined,
  };

  /**
   * Map of event handlers to be applied to the embedded report
   */
  // Update event handlers for the report by redefining the map using this.eventHandlersMap
  // Set event handler to null if event needs to be removed
  // More events can be provided from here
  // https://docs.microsoft.com/en-us/javascript/api/overview/powerbi/handle-events#report-events
  eventHandlersMap = new Map ([
    ['loaded', () => {
        const report = this.reportObj.getReport();
        report.setComponentTitle('Embedded report');
        console.log('Report has loaded');
      },
    ],
    ['rendered', () => console.log('Report has rendered')],
    ['error', (event?: service.ICustomEvent<any>) => {
        if (event) {
          console.error(event.detail);
        }
      },
    ],
    ['visualClicked', () => console.log('visual clicked')],
    ['pageChanged', (event) => console.log(event)],
  ]) as Map<string, (event?: service.ICustomEvent<any>, embeddedEntity?: Embed) => void | null>;

  constructor(public httpService: HttpService, private element: ElementRef<HTMLDivElement>) {}

  /**
   * Embeds report
   *
   * @returns Promise<void>
   */
  async embedReport(): Promise<void> {
    let reportConfigResponse: ConfigResponse;

    // Get the embed config from the service and set the reportConfigResponse
    try {
      reportConfigResponse = {
        "embedReports": 
          {
            "embedUrl": "https://app.powerbi.com/reportEmbed",
          },
        "Id": "7e115028-236b-4d59-a99a-7abb8aeba921",
        "embedToken": "H4sIAAAAAAAEAB3Tt66sZgAE4Hc5LZaAZUmWbkHOLDl15LCEnxwsv7uP3I-m-Ebzz4-V3v2UFj9__1jFOtsiey2E5CnMg_JGPcg-Fg0OaW5t2y9VEIDmMwaI5J-ymfBtc52lDNtXc2uVhxZ7Ib8kwsGw5rNJY3i0GX8MCfdZrCIQQpwfu8tyHrIxqvrlqy6pUbz88oSMrQuX3KeLz2cVZw45oAMvT75lws6Yk55yFhI8g_sm823bx0bAer_IxY5ibHGzkDxGyMseTrNbgxBKiPeKi2w6I4N3j8LvDwI8eyvB8mGpPvuWr_wbsR2RYFwzYlOjS1PR4he42hlrP08oYj1NeQhmLXKqnMmQAPaZli4LeRyl7gzFs2zo96fpwRsAMXSQuJi3vhh3Ph-pHKpoRZbtii51b2AfYocrl_PB9w6fVkGqWDrY9Ez3B8hi0gtuj0XyghJnvJ4AQV0_H0nrZUeTFxIMgyfqi1dac-ErU72x3iFofGnEU0Hl7LuE-S9P7GJKnnrymfHiOBvBAqIYRx9auaybJrz5YEdO0mWDcqDBSbyd5xSMj4SVTu0reTiPj-PgWpaujtiyMNRHipcIA7Up6V-HA3penX7ZlSOZdEyVpXzE6xJT0i9QmD2vEBk1pCFrZAjl2y0tvHyB5D7d1C74_aZL18HtwDf9LIZ3KijxFZ5VqH8HarAGu6Q8Dzsb_AdR-d1XmK1vlGVwbzhkzlUCq4s0G3lJx60JUy0fGzhh__HBPFxgirv4YLD23FCzKhPUxYfmzWhuF1KPG6GrI_IwjtrJ_q0Krwxq-7qcL99RFvaB0NycJfQt-_xwe1fNXn7SCiwpcMw9-oTV60dmj3zX9wq5UZCmvHkTa-G8XLvpFS0DqJpK9qUVOpU_P3_9cMsNtkkr79_rJE6Ye6Ivwyl5A7vAJfOczgeYKkHWwKGnWz01yIhg1Fuso33q8SvaOSm16bqaNR99pr4DyjwaR1wj5e2mY35bYjP3F9wQA6Opj4P25lEnF5w9ondIHyKUcxD1W-jcA64sJsPSeJSJBQvOMz-2zty_sCtdqg2lVpjttOervUSdj_btY-yGhGL8At2oKVflQoEh-CKazLV8-xPheMc0GXzMSfNv91UER1itFhPFFsbXH2YPUFnPSycIUp-ECBiFgsRwOhphyurSC9bAkilW_SPEfNcusTo0MA0gkYL5lWi_ROyckix-U7oTKlPWqdAWVAheu59-3E1ds9xYFaJ7dohgpM4__zPfoCkXJfhVttmsZUxKycrDVWh0nOOvQ9v_p9y2HtNtX8rf2GZG8PbhpmyYfxcy37WUxE2TTLJJ9GM3OOKycbSmFV-tQStAN8a90KCSBNtXEYwMYDp7Wtyr-w9W2vVRsbfaiDImeCYpiOc6q7xdn43TjXSSZ4QZNNpCFzbeHFyjI2LEhYxs-iVNrOYTr-GnYWHjguSGGP2FMW5I_2jacatZub9TjCwQApK7lwscsdpV-xKRKzq1Sm3bpoW3vKIYJyHnUnxTIIbyWk90SGhrq7b0a4PaAquTzkf1dDmqMSVPDjkWN0TeXfV23rHtFXYmEtWW6sR7U0c1iF6C2g88VgZuaVq16Y7m4XthqK2eawmRT0htEoScPYDka8aTeFIeW5vX_k1G5pf53_8Anu_SdS4GAAA=.eyJjbHVzdGVyVXJsIjoiaHR0cHM6Ly9ERi1NU0lULVNDVVMtcmVkaXJlY3QuYW5hbHlzaXMud2luZG93cy5uZXQiLCJleHAiOjE3Mjk3NTY3MDQsImFsbG93QWNjZXNzT3ZlclB1YmxpY0ludGVybmV0Ijp0cnVlfQ=="
      };
      
      //await this.httpService.getEmbedConfig(reportUrl).toPromise();
    } catch (error: any) {
      this.displayMessage = `Failed to fetch config for report. Status: ${error.status} ${error.statusText}`;
      console.error(this.displayMessage);
      return;
    }

    // Update the reportConfig to embed the PowerBI report
    this.reportConfig = {
      ...this.reportConfig,
      id: reportConfigResponse.Id,
      embedUrl: reportConfigResponse.embedReports.embedUrl ,
      accessToken: reportConfigResponse.embedToken,
    };

    // Update embed status
    this.isEmbedded = true;

    // Update the display message
    this.displayMessage = 'Use the buttons above to interact with the report using Power BI Client APIs.';
  }

  /**
   * Change Visual type
   *
   * @returns Promise<void>
   */
  async changeVisualType(): Promise<void> {
    // Get report from the wrapper component
    const report: Report = this.reportObj.getReport();

    if (!report) {
      this.displayMessage = 'Report not available.';
      console.log(this.displayMessage);
      return;
    }

    // Get all the pages of the report
    const pages: Page[] = await report.getPages();

    // Check if the pages are available
    if (pages.length === 0) {
      this.displayMessage = 'No pages found.';
      return;
    }

    // Get active page of the report
    const activePage: Page | undefined = pages.find((page) => page.isActive);

    if (!activePage) {
      this.displayMessage = 'No Active page found';
      return;
    }

    try {
      // Change the visual type using powerbi-report-authoring
      // For more information: https://docs.microsoft.com/en-us/javascript/api/overview/powerbi/report-authoring-overview
      // Get the visual
      const visual = await activePage.getVisualByName('VisualContainer6');

      const response = await visual.changeType('lineChart');

      this.displayMessage = `The ${visual.type} was updated to lineChart.`;

      console.log(this.displayMessage);

      return response;
    } catch (error) {
      if (error === 'PowerBIEntityNotFound') {
        console.log('No Visual found with that name');
      } else {
        console.log(error);
      }
    }
  }

  /**
   * Hide Filter Pane
   *
   * @returns Promise<IHttpPostMessageResponse<void> | undefined>
   */
  async hideFilterPane(): Promise<IHttpPostMessageResponse<void> | undefined> {
    // Get report from the wrapper component
    const report: Report = this.reportObj.getReport();

    if (!report) {
      this.displayMessage = 'Report not available.';
      console.log(this.displayMessage);
      return;
    }

    // New settings to hide filter pane
    const settings = {
      panes: {
        filters: {
          expanded: false,
          visible: false,
        },
      },
    };

    try {
      const response = await report.updateSettings(settings);
      this.displayMessage = 'Filter pane is hidden.';
      console.log(this.displayMessage);

      return response;
    } catch (error) {
      console.error(error);
      return;
    }
  }

  /**
   * Set data selected event
   *
   * @returns void
   */
  setDataSelectedEvent(): void {
    // Adding dataSelected event in eventHandlersMap
    this.eventHandlersMap = new Map<string, (event?: service.ICustomEvent<any>, embeddedEntity?: Embed) => void | null>([
      ...this.eventHandlersMap,
      ['dataSelected', (event) => console.log(event)],
    ]);

    this.displayMessage = 'Data Selected event set successfully. Select data to see event in console.';
  }

  onRegionSelected(value:string){
    console.log("the selected value is " + value);
    const report: any = this.reportObj.getReport();
    report.setFilters([{
      $schema: "",
      target: {
        table: "Region",
        column: "Region",
      },
      values: [value],
      operator: 'In',
      filterType: 1,
    }])
 }
 onCatSelected(value:string){
  console.log("the selected value is " + value);
    const report: any = this.reportObj.getReport();
    report.savePersistentFilters();
    report.setFilters([{
      $schema: "",
      target: {
        table: "Product",
        column: "category",
      },
      values: [value],
      operator: 'In',
      filterType: 1,
    }])
}

}
