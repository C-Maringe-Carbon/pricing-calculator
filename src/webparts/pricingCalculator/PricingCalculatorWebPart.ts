// @ts-nocheck
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { initCalculator } from './calculator';

export default class ServicesCalculatorWebPart extends BaseClientSideWebPart<{}> {

  public render(): void {
    initCalculator(this.domElement, this.context.pageContext.web.absoluteUrl);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration() {
    return { pages: [] };
  }
}
