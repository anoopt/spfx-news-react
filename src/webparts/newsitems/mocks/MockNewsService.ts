import { INewsItem } from '../interfaces/INewsItem'
import { INewsItems } from '../interfaces/INewsItems'
import { INewsService } from '../interfaces/INewsService'

import { ServiceScope } from '@microsoft/sp-client-base';

export class MockNewsService implements INewsService {
  constructor(serviceScope: ServiceScope) {
  }

  public loadNewsItemsUsingService(siteUrl: string, numberOfItems: number, listName: string) : Promise<INewsItems>{
    return new Promise<INewsItems>((resolve, reject) => {
      const newsitems: INewsItems = { value :
      [
        {
          Id: '1',
          Title: 'Mock News 1',
          ImageUrl: {
            Url: 'http://www.stmichaelschurchwatersupton.org.uk/Location-News-icon.png'
          },
          Byline:'This is mock news 1 from service'
        },
        {
          Id: '2',
          Title: 'Mock News 2',
          ImageUrl: {
            Url: 'http://www.stmichaelschurchwatersupton.org.uk/Location-News-icon.png'
          },
          Byline:'This is mock news 2 from service'
        },
        {
          Id: '3',
          Title: 'Mock News 3',
          ImageUrl: {
            Url: 'http://www.stmichaelschurchwatersupton.org.uk/Location-News-icon.png'
          },
          Byline:'This is mock news 3 from service'
        },
        {
          Id: '4',
          Title: 'Mock News 4',
          ImageUrl: {
            Url: 'http://www.stmichaelschurchwatersupton.org.uk/Location-News-icon.png'
          },
          Byline:'This is mock news 4 from service'
        }
      ]
      };

      var retItems: INewsItems = {
      value :
      [
        {
          Id: '1',
          Title: 'Mock News 1',
          ImageUrl: {
            Url: 'http://www.stmichaelschurchwatersupton.org.uk/Location-News-icon.png'
          },
          Byline:'This is mock news 1 from service'
        }
      ]
      };

      newsitems.value.map((item: INewsItem, i: number) => {
        if(i < numberOfItems){
          retItems.value[i] = item;
        }
      });
      resolve(retItems);
    });
  }

  public loadNewsItemsUsingPnPService(siteUrl: string, numberOfItems: number, listName: string) : Promise<INewsItem[]>{
    return new Promise<INewsItem[]>((resolve, reject) => {
      const newsitems: INewsItems = { value :
      [
        {
          Id: '1',
          Title: 'Mock News 1',
          ImageUrl: {
            Url: 'http://www.stmichaelschurchwatersupton.org.uk/Location-News-icon.png'
          },
          Byline:'This is mock news 1 from service'
        },
        {
          Id: '2',
          Title: 'Mock News 2',
          ImageUrl: {
            Url: 'http://www.stmichaelschurchwatersupton.org.uk/Location-News-icon.png'
          },
          Byline:'This is mock news 2 from service'
        },
        {
          Id: '3',
          Title: 'Mock News 3',
          ImageUrl: {
            Url: 'http://www.stmichaelschurchwatersupton.org.uk/Location-News-icon.png'
          },
          Byline:'This is mock news 3 from service'
        },
        {
          Id: '4',
          Title: 'Mock News 4',
          ImageUrl: {
            Url: 'http://www.stmichaelschurchwatersupton.org.uk/Location-News-icon.png'
          },
          Byline:'This is mock news 4 from service'
        }
      ]
      };

      var retItems: INewsItem[] =
      [
        {
          Id: '1',
          Title: 'Mock News 1',
          ImageUrl: {
            Url: 'http://www.stmichaelschurchwatersupton.org.uk/Location-News-icon.png'
          },
          Byline:'This is mock news 1 from service'
        }
      ];

      newsitems.value.map((item: INewsItem, i: number) => {
        if(i < numberOfItems){
          retItems[i] = item;
        }
      });
      resolve(retItems);
    });
  }
}