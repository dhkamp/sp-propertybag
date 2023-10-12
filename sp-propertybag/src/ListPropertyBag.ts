import { SPFI, SPFx, spfi } from "@pnp/sp";
import { IList } from "@pnp/sp/lists";
import { BaseComponentContext } from "@microsoft/sp-component-base";
import { stringIsNullOrEmpty } from "@pnp/core";
import { IFolderInfo } from "@pnp/sp/folders";
import { SPHttpClient } from "@microsoft/sp-http-base";

import "@pnp/sp/sites";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/folders";

export class ListPropertyBag {
  private readonly sp: SPFI;
  private readonly list: IList;
  private readonly context: BaseComponentContext;

  constructor(context: BaseComponentContext, listId: string) {
    this.sp = spfi().using(SPFx(context));
    this.context = context;
    this.list = this.sp.web.lists.getById(listId);
  }
  public async addOrUpdate(key: string, value: string) {
    const web = await this.sp.web.select("Id")();
    const site = await this.sp.site.select("Id")();
    const rootFolder = await this.list.rootFolder();

    const payload = `
    <Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="MyApplicationName">
      <Actions>
        <Method Name="SetFieldValue" Id="9" ObjectPathId="4">
          <Parameters>
            <Parameter Type="String">${key}</Parameter>
            <Parameter Type="String">${value}</Parameter>
          </Parameters>
        </Method>
        <Method Name="Update" Id="10" ObjectPathId="2" />
      </Actions>
      <ObjectPaths>
        <Identity Id="2" Name="740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${site.Id}:web:${web.Id}:folder:${rootFolder.UniqueId}" />
        <Property Id="4" ParentId="2" Name="Properties" />
      </ObjectPaths>
    </Request>`;

    const client = this.context.spHttpClient;
    client.post(
      `${web.Url}/_vti_bin/client.svc/ProcessQuery`,
      SPHttpClient.configurations.v1,
      {
        method: "POST",
        headers: {
          Accept: "*/*",
          "Content-Type": 'text/xml;charset="UTF-8"',
          "X-Requested-With": "XMLHttpRequest",
        },
        body: payload,
      },
    );
  }
  public async get(key: string) {
    type RootFolderWithPropertyBag = IFolderInfo & {
      Properties: { [key: string]: string };
    };

    const list = await this.list.expand("RootFolder/Properties")();

    const properties = (list.RootFolder as RootFolderWithPropertyBag)
      .Properties;

    return properties[key];
  }

  /**
   * Checks if the specified key exists in the property bag and has the specified value.
   * If no value is provided, it checks if the key exists and has a non-empty value.
   * @param key - The key to check in the property bag.
   * @param value - The value to compare against the value in the property bag.
   * @returns True if the key exists and has the specified value, or if no value is provided and the key exists with a non-empty value. False otherwise.
   */
  public async has(key: string, value?: string): Promise<boolean> {
    const propertyBagValue = await this.get(key);
    return stringIsNullOrEmpty(value)
      ? stringIsNullOrEmpty(propertyBagValue)
      : propertyBagValue === value;
  }
}
