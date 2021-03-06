import { INewsItem } from '../interfaces/INewsItem';

export default class MockHttpClient {

  private static _items: INewsItem[] = [
    {
      Id: '1',
      Title: 'Mock News 1',
      ImageUrl: {
        Url: 'http://www.bbc.co.uk/news/special/2015/newsspec_10857/bbc_news_logo.png'
      },
      ProfileImageUrl: {
        Url: 'http://www.stmichaelschurchwatersupton.org.uk/Location-News-icon.png'
      },
      Byline:'This is mock news 1',
      Author:{
        Title:'Anoop'
      }
    },
    {
      Id: '2',
      Title: 'Mock News 2',
      ImageUrl: {
        Url: 'http://www.bbc.co.uk/news/special/2015/newsspec_10857/bbc_news_logo.png'
      },
      ProfileImageUrl: {
        Url: 'http://www.stmichaelschurchwatersupton.org.uk/Location-News-icon.png'
      },
      Byline:'This is mock news 2',
      Author:{
        Title:'Anoop'
      }
    },
    ];

  public static get(restUrl: string, options?: any): Promise<INewsItem[]> {
    return new Promise<INewsItem[]>((resolve) => {
      resolve(MockHttpClient._items);
    });
  }

  public static getItems(restUrl: string, numberOfItems: number, options?: any): Promise<INewsItem[]> {

    var items: INewsItem[] = [
      {
        Id: '1',
        Title: 'Mock News 1',
        ImageUrl: {
          Url: 'http://www.bbc.co.uk/news/special/2015/newsspec_10857/bbc_news_logo.png'
        },
        ProfileImageUrl: {
          Url: 'http://www.stmichaelschurchwatersupton.org.uk/Location-News-icon.png'
        },
        Byline:'This is mock news 1',
      Author:{
        Title:'Anoop'
      }
      },
      {
        Id: '2',
        Title: 'Mock News 2',
        ImageUrl: {
          Url: 'http://www.bbc.co.uk/news/special/2015/newsspec_10857/bbc_news_logo.png'
        },
        ProfileImageUrl: {
          Url: 'http://www.stmichaelschurchwatersupton.org.uk/Location-News-icon.png'
        },
        Byline:'This is mock news 2',
      Author:{
        Title:'Anoop'
      }
      },
      {
        Id: '3',
        Title: 'Mock News 3',
        ImageUrl: {
          Url: 'http://www.bbc.co.uk/news/special/2015/newsspec_10857/bbc_news_logo.png'
        },
        ProfileImageUrl: {
          Url: 'http://www.stmichaelschurchwatersupton.org.uk/Location-News-icon.png'
        },
        Byline:'This is mock news 3',
      Author:{
        Title:'Anoop'
      }
      },
      {
        Id: '4',
        Title: 'Mock News 4',
        ImageUrl: {
          Url: 'http://www.bbc.co.uk/news/special/2015/newsspec_10857/bbc_news_logo.png'
        },
        ProfileImageUrl: {
          Url: 'http://www.stmichaelschurchwatersupton.org.uk/Location-News-icon.png'
        },
        Byline:'This is mock news 4',
      Author:{
        Title:'Anoop'
      }
      }
    ];

    var retItems: INewsItem[] = [];

    items.map((item: INewsItem, i: number) => {
      if(i < numberOfItems){
        retItems.push(item);
      }
    });

    return new Promise<INewsItem[]>((resolve) => {
      resolve(retItems);
    });
  }
}