import { INewsItem } from '../interfaces/INewsItem'
import { INewsItems } from '../interfaces/INewsItems'

export interface INewsService {
  loadNewsItemsUsingService: (siteUrl: string, numberOfItems: number, listName: string) => Promise<INewsItems>;
  loadNewsItemsUsingPnPService: (siteUrl: string, numberOfItems: number, listName: string) => Promise<INewsItem[]>;
}