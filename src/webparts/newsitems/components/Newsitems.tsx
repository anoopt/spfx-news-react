import * as React from 'react';

import styles from '../Newsitems.module.scss';
import { INewsitemsWebPartProps } from '../INewsitemsWebPartProps';
import { HttpClient, EnvironmentType, ServiceScope, ServiceKey } from '@microsoft/sp-client-base';
import { SearchUtils, ISearchQueryResponse, IRow } from '../../SearchUtils';
import {
  css,
  Persona,
  PersonaSize,
  PersonaPresence,
  DocumentCard,
  DocumentCardPreview,
  DocumentCardTitle,
  DocumentCardActivity,
  Spinner
} from 'office-ui-fabric-react';
import { INewsItem } from '../interfaces/INewsItem'
import { INewsItems } from '../interfaces/INewsItems'
import { INewsService } from '../interfaces/INewsService'
import MockHttpClient from '../mocks/MockHttpClient';
import { MockNewsService } from '../mocks/MockNewsService';
import { NewsService } from '../services/NewsService';


export interface INewsitemsProps extends INewsitemsWebPartProps {
  httpClient: HttpClient;
  siteUrl: string;
  listName: string;
  environmentType: EnvironmentType;
  serviceScope: ServiceScope;
}

export interface INewsItemState {
  newsItems: INewsItem[];
  loading: boolean;
  error: string;
}

export default class Newsitems extends React.Component<INewsitemsProps, INewsItemState> {

  private newsServiceInstance: INewsService;

  constructor(props: INewsitemsProps, state: INewsItemState) {
    super(props);

    this.state = {
      newsItems: [] as INewsItem[],
      loading: true,
      error: null
    };

    let serviceScope: ServiceScope;
    const newsServiceKey: ServiceKey<INewsService> = ServiceKey.create<INewsService>("newsServiceKey", NewsService);

    const currentEnvType = this.props.environmentType;
    if (currentEnvType == EnvironmentType.SharePoint || currentEnvType == EnvironmentType.ClassicSharePoint) {

      serviceScope = this.props.serviceScope;

    }
    else {

      serviceScope = this.props.serviceScope.startNewChild();
      serviceScope.createAndProvide(newsServiceKey, MockNewsService);
      serviceScope.finish();
    }

    this.newsServiceInstance = serviceScope.consume(newsServiceKey);
  }

   private _getMockListData(): Promise<INewsItems> {
    return MockHttpClient.getItems(this.props.siteUrl, this.props.numberOfItems)
            .then((data: INewsItem[]) => {
                 var listData: INewsItems = { value: data };
                 return listData;
             }) as Promise<INewsItems>;
  }

  public componentDidMount(): void {
    //this.loadNewsItemsNormal(this.props.siteUrl, this.props.numberOfItems, this.props.listName);
    //this.loadNewsItemsFromService(this.props.siteUrl, this.props.numberOfItems, this.props.listName);
    this.loadNewsItemsFromPnPService(this.props.siteUrl, this.props.numberOfItems, this.props.listName);

  }

  public componentDidUpdate(prevProps: INewsitemsProps, prevState: INewsItemState, prevContext: any): void {
    if (this.props.numberOfItems !== prevProps.numberOfItems ||
      this.props.siteUrl !== prevProps.siteUrl && (
        this.props.numberOfItems && this.props.siteUrl
      )) {
      //this.loadNewsItemsNormal(this.props.siteUrl, this.props.numberOfItems, this.props.listName);
      //this.loadNewsItemsFromService(this.props.siteUrl, this.props.numberOfItems, this.props.listName);
      this.loadNewsItemsFromPnPService(this.props.siteUrl, this.props.numberOfItems, this.props.listName);
    }
  }


  public render(): JSX.Element {
    const loading: JSX.Element = this.state.loading ? <div style={{ margin: '0 auto' }}><Spinner label={'Loading...'} /></div> : <div/>;
    const error: JSX.Element = this.state.error ? <div><strong>Error: </strong> {this.state.error}</div> : <div/>;
    const newsItems: JSX.Element[] = this.state.newsItems.map((newsitem: INewsItem, i: number) => {

      return (
        //<Persona
          //primaryText={newsitem.Title}
          //secondaryText={newsitem.Byline}
          //imageUrl={newsitem.ImageUrl.Url}
        ///>

         <DocumentCard key={newsitem.Id}>
          <DocumentCardPreview
            previewImages={[
              {
                previewImageSrc: newsitem.ImageUrl.Url,
                width: 318,
                height: 196,
                accentColor: '#ce4b1f'
              }
            ]}
            />
          <DocumentCardTitle title={newsitem.Title} />
          <DocumentCardActivity
            activity={newsitem.Author.Title}
            people={
              [
                { name: newsitem.Byline, profileImageSrc: newsitem.ProfileImageUrl.Url }
              ]
            }
            />
        </DocumentCard>
      );
    });

    return (
      <div className={styles.newsItems}>

        {loading}
        {error}
        {newsItems}
      </div>
    );
  }


  private navigateTo(url: string): void {
    window.open(url, '_blank');
  }

  private loadNewsItemsFromPnPService(siteUrl: string, numberOfItems: number, listName: string): void{
    this.newsServiceInstance.loadNewsItemsUsingPnPService(siteUrl, numberOfItems, listName).
    then((newsItemRet: INewsItem[] ): void => {
      console.log(newsItemRet);
          this.setState({
            loading: false,
            error: null,
            newsItems: newsItemRet
          });
        }, (error: any): void => {
          this.setState({
            loading: false,
            error: error,
            newsItems: []
          });
        });
  }

  private loadNewsItemsFromService(siteUrl: string, numberOfItems: number, listName: string): void{
    this.newsServiceInstance.loadNewsItemsUsingService(siteUrl, numberOfItems, listName).
    then((newsItemRet: { value: INewsItem[] }): void => {
      //console.log(newsItemRet);
          this.setState({
            loading: false,
            error: null,
            newsItems: newsItemRet.value
          });
        }, (error: any): void => {
          this.setState({
            loading: false,
            error: error,
            newsItems: []
          });
        });
  }

  private loadNewsItemsNormal(siteUrl: string, numberOfPeople: number, listName: string): void{
    const currentEnvType = this.props.environmentType;

    if (currentEnvType == EnvironmentType.SharePoint || currentEnvType == EnvironmentType.ClassicSharePoint) {

      this.props.httpClient.get(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Title,Id,ImageUrl,ProfileImageUrl,Byline&$top=${numberOfPeople}`, {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response: Response): Promise<{ value: INewsItem[] }> => {
          return response.json();
        })
        .then((response: { value: INewsItem[] }): void => {
          console.log(response.value);
          this.setState({
            loading: false,
            error: null,
            newsItems: response.value
          });
        }, (error: any): void => {
          this.setState({
            loading: false,
            error: error,
            newsItems: []
          });
        });
    }
    else {
      this._getMockListData().then((response: { value: INewsItem[] }): void => {
        this.setState({
          loading: false,
          error: null,
          newsItems: response.value
        });
      }, (error: any): void => {
        this.setState({
          loading: false,
          error: error,
          newsItems: []
        });
      });
    }
  }

   private loadNewsItemsUsingSearch(siteUrl: string, numberOfPeople: number): void {
    this.props.httpClient.get(`${siteUrl}/_api/search/query?querytext='path:${siteUrl}/Lists/News%20contentclass:"STS_ListItem"'&selectproperties='Title,Id'&rowlimit=${numberOfPeople}`, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'odata-version': ''
      }
    })
      .then((response: Response): Promise<ISearchQueryResponse> => {
        return response.json();
      })
      .then((response: ISearchQueryResponse): void => {
        if (!response ||
          !response.PrimaryQueryResult ||
          !response.PrimaryQueryResult.RelevantResults ||
          response.PrimaryQueryResult.RelevantResults.RowCount === 0) {
          this.setState({
            loading: false,
            error: null,
            newsItems: []
          });
          return;
        }

        const newsItems: INewsItem[] = [];
        for (let i: number = 0; i < response.PrimaryQueryResult.RelevantResults.Table.Rows.length; i++) {
          const newsRow: IRow = response.PrimaryQueryResult.RelevantResults.Table.Rows[i];

          newsItems.push({
            Title: SearchUtils.getValueFromResults('Title', newsRow.Cells),
            Id: SearchUtils.getValueFromResults('Id', newsRow.Cells),
            ImageUrl: "",
            Byline:"",
            ProfileImageUrl: "",
            Author: ""
          });
        }

        this.setState({
          loading: false,
          error: null,
          newsItems: newsItems
        });
      }, (error: any): void => {
        this.setState({
          loading: false,
          error: error,
          newsItems: []
        });
      });
  }
}
