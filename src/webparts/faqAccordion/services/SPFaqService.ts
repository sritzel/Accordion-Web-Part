import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';
import { IFaqItem } from '../models/IFaqItem';

interface ISPFaqListItem {
  Id: number;
  Title?: string;           // used if Question column is the default Title column
  Question?: string;
  Answer?: string;
  Category?: string;
  SortOrder?: number;
}

interface ISPItemsResponse {
  value: ISPFaqListItem[];
  'odata.nextLink'?: string;
}

export class SPFaqService {

  private readonly context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  /**
   * Fetch all FAQ items from the given list id.
   * Pages through results if the list has more than 5000 items.
   *
   * @param listId GUID of the SharePoint list
   * @param siteUrl optional absolute URL of the site that contains the list.
   *                When omitted (or empty), uses the current site.
   * @param setId  optional Set lookup ID. When provided, items are filtered
   *               server-side by `SetId eq <setId>`. The lookup column is
   *               assumed to have internal name "Set" — change SET_COLUMN
   *               below if your column is named differently.
   */
  public async getFaqItems(listId: string, siteUrl?: string, setId?: string): Promise<IFaqItem[]> {
    if (!listId) {
      return [];
    }

    const baseUrl = (siteUrl && siteUrl.trim())
      ? siteUrl.replace(/\/+$/, '')
      : this.context.pageContext.web.absoluteUrl;

    const SET_COLUMN = 'Set'; // internal name of the lookup column

    // Select both Title and Question to cover the case where Question is either
    // a renamed Title column or a separate column.
    const select = '$select=Id,Title,Question,Answer,Category,SortOrder';
    const top = '$top=4999';
    let initialUrl =
      `${baseUrl}/_api/web/lists(guid'${listId}')/items?${select}&${top}`;

    if (setId && setId.trim() !== '') {
      // For a lookup column named "Set", the ID is exposed as "SetId" in REST.
      initialUrl += `&$filter=${SET_COLUMN}Id eq ${encodeURIComponent(setId)}`;
    }

    const all: ISPFaqListItem[] = [];
    let nextUrl: string | undefined = initialUrl;

    while (nextUrl) {
      const page: ISPItemsResponse = await this._getJson(nextUrl);
      all.push(...page.value);
      nextUrl = page['odata.nextLink'];
    }

    return all.map(this._mapItem).filter(i => !!i.question);
  }

  /**
   * Read the schema of a lookup column and return the GUID of the list it
   * points to. Used to auto-detect the "FAQ Sets" list from the "Set" column
   * on the FAQ list, so the user doesn't have to pick it manually.
   *
   * Returns an empty string if the column doesn't exist or isn't a lookup.
   */
  public async getLookupTargetListId(
    faqListId: string,
    columnInternalName: string,
    siteUrl?: string
  ): Promise<string> {
    if (!faqListId || !columnInternalName) return '';
    const base = (siteUrl && siteUrl.trim())
      ? siteUrl.replace(/\/+$/, '')
      : this.context.pageContext.web.absoluteUrl;

    const url =
      `${base}/_api/web/lists(guid'${faqListId}')` +
      `/fields/getbyinternalnameortitle('${encodeURIComponent(columnInternalName)}')` +
      `?$select=LookupList,TypeAsString`;

    const options: ISPHttpClientOptions = {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'odata-version': ''
      }
    };
    const res: SPHttpClientResponse = await this.context.spHttpClient.get(
      url, SPHttpClient.configurations.v1, options
    );
    if (!res.ok) {
      // Column doesn't exist or no permission — return empty so the caller
      // can degrade gracefully.
      return '';
    }
    const data = await res.json() as { LookupList?: string; TypeAsString?: string };
    // LookupList is reported as "{GUID}" — strip braces for our REST URLs.
    const raw = (data && data.LookupList) || '';
    return raw.replace(/^\{|\}$/g, '');
  }

  /**
   * Fetch a simple list of {id, title} items from any list — used to populate
   * the "Filter by Set" dropdown from the FAQ Sets list.
   */
  public async getListChoices(listId: string, siteUrl?: string): Promise<{ id: number; title: string }[]> {
    if (!listId) return [];
    const base = (siteUrl && siteUrl.trim())
      ? siteUrl.replace(/\/+$/, '')
      : this.context.pageContext.web.absoluteUrl;

    const url = `${base}/_api/web/lists(guid'${listId}')/items?$select=Id,Title&$top=4999&$orderby=Title%20asc`;
    const out: { id: number; title: string }[] = [];
    let next: string | undefined = url;
    while (next) {
      const page: ISPItemsResponse = await this._getJson(next);
      for (let i = 0; i < page.value.length; i++) {
        const v = page.value[i];
        out.push({ id: v.Id, title: (v.Title || '').trim() });
      }
      next = page['odata.nextLink'];
    }
    return out;
  }

  private _mapItem = (raw: ISPFaqListItem): IFaqItem => {
    const rawSort = raw.SortOrder;
    const sortOrder = (typeof rawSort === 'number' && !isNaN(rawSort)) ? rawSort : undefined;
    return {
      id: raw.Id,
      question: (raw.Question || raw.Title || '').trim(),
      answer: raw.Answer || '',
      category: (raw.Category || 'Uncategorized').trim(),
      sortOrder: sortOrder
    };
  };

  private async _getJson(url: string): Promise<ISPItemsResponse> {
    const options: ISPHttpClientOptions = {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'odata-version': ''
      }
    };
    const res: SPHttpClientResponse = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      options
    );
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`SharePoint request failed (${res.status}): ${text}`);
    }
    return (await res.json()) as ISPItemsResponse;
  }
}
