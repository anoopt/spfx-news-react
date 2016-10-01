import { INewsItem } from '../interfaces/INewsItem'
import { INewsItems } from '../interfaces/INewsItems'
import { INewsService } from '../interfaces/INewsService'

import { ServiceScope, HttpClient, IODataBatchOptions, ODataBatch, httpClientServiceKey } from '@microsoft/sp-client-base';

import * as pnp from 'sp-pnp-js';

export class NewsService implements INewsService {
  private httpClient: HttpClient;

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {
      this.httpClient = serviceScope.consume(httpClientServiceKey);
    });

    pnp.setup({
      headers: {
        'Accept': 'application/json;odata=nometadata'
      }
    });
  }

  public loadNewsItemsUsingPnPService(siteUrl: string, numberOfItems: number, listName: string) : Promise<INewsItem[]>{
    return pnp.sp.web.lists.getByTitle(listName)
      .items.select('Title', 'Id', 'ImageUrl','Byline').
      top(numberOfItems).
      get();
  }

  public loadNewsItemsUsingService(siteUrl: string, numberOfItems: number, listName: string) : Promise<INewsItems>{
    return this.httpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Title,Id,ImageUrl,Byline&$top=${numberOfItems}`,
      {
        headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
      })
      .then((response: Response) => {
        console.log(response);
        return response.json();
      });
  }
}