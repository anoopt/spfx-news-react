import * as React from 'react';

import styles from '../Newsitems.module.scss';
import { INewsitemsWebPartProps } from '../INewsitemsWebPartProps';
import { HttpClient } from '@microsoft/sp-client-base';
import { SearchUtils, ISearchQueryResponse, IRow } from '../../SearchUtils';
import {
  css,
  Persona,
  PersonaSize,
  PersonaPresence,
  Spinner
} from 'office-ui-fabric-react';

export interface INewsitemsProps extends INewsitemsWebPartProps {
  httpClient: HttpClient;
  siteUrl: string;
}

export interface INewsItemState {
  newsItems: INewsItem[];
  loading: boolean;
  error: string;
}

export interface INewsItem {
  id: string;
  title: string;
  imageUrl: string;
}

interface ISearchResultValue {
  Key: string;
  Value: string;
}


export default class Newsitems extends React.Component<INewsitemsProps, INewsItemState> {

  constructor(props: INewsitemsProps, state: INewsItemState) {
    super(props);

    this.state = {
      newsItems: [] as INewsItem[],
      loading: true,
      error: null
    };
  }

  public componentDidMount(): void {
    this.loadNewsItems(this.props.siteUrl, this.props.numberOfItems);
  }

  public componentDidUpdate(prevProps: INewsitemsProps, prevState: INewsItemState, prevContext: any): void {
    if (this.props.numberOfItems !== prevProps.numberOfItems ||
      this.props.siteUrl !== prevProps.siteUrl && (
        this.props.numberOfItems && this.props.siteUrl
      )) {
      this.loadNewsItems(this.props.siteUrl, this.props.numberOfItems);
    }
  }


  public render(): JSX.Element {
    const loading: JSX.Element = this.state.loading ? <div style={{ margin: '0 auto' }}><Spinner label={'Loading...'} /></div> : <div/>;
    const error: JSX.Element = this.state.error ? <div><strong>Error: </strong> {this.state.error}</div> : <div/>;
    const newsItems: JSX.Element[] = this.state.newsItems.map((newsitem: INewsItem, i: number) => {
      return (
        <Persona
          primaryText={newsitem.title}
          secondaryText={newsitem.id}
        />
      );
    });

    return (
      <div className={styles.workingWith}>
        <div className={css('ms-font-xl', styles.webPartTitle)}>{this.props.title}</div>
        {loading}
        {error}
        {newsItems}
      </div>
    );
  }


   private navigateTo(url: string): void {
    window.open(url, '_blank');
  }

   private loadNewsItems(siteUrl: string, numberOfPeople: number): void {
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
            title: SearchUtils.getValueFromResults('Title', newsRow.Cells),
            id: SearchUtils.getValueFromResults('Id', newsRow.Cells),
            imageUrl: ""
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
