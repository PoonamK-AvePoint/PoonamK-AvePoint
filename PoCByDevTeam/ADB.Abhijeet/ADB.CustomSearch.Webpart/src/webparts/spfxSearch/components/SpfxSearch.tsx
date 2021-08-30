import * as React from 'react';

import styles from './SpfxSearch.module.scss';
import { ISpfxSearchProps } from './ISpfxSearchProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { SearchBox, Dropdown, IDropdownOption, DropdownMenuItemType, PrimaryButton, DefaultButton, IIconProps, IChoiceGroupOption, StackItem } from '@fluentui/react';
import { Link, Text, Label } from '@fluentui/react';
import { Stack, IStackTokens } from '@fluentui/react';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn } from '@fluentui/react/lib/DetailsList';
import { SearchResult, SearchResults, sp } from 'sp-pnp-js';
import { ISearchResult, IRefiners, IDetailsListSearchResultState, IRefinerFilter } from './ISearchResult';
import { RoleAssignments } from 'sp-pnp-js/lib/sharepoint/roles';
import ExcelExport from 'export-xlsx';
import { SETTINGS_FOR_EXPORT, SETTINGS_RESULT_EXPORT } from './settings';
import Pagination from "react-js-pagination";

//import Pagination from 'react-bootstrap-4-pagination';
import 'bootstrap/dist/css/bootstrap.min.css';
import { filter } from 'lodash';
//import * as strings from 'SpfxSearchWebPartStrings';
var arrResult = [];

export default class SpfxSearch extends React.Component<ISpfxSearchProps, { searchText: string, results: ISearchResult[], activePage: Number, refinersResult: IRefiners[], chociceOption?: any[], pnpResults?: SearchResults, totalRowsCount: Number, seletedprojects: any, concatRefiners?: any, refiners:any[],refinementFilters:IRefinerFilter[] }> {

  private currentPage: number = 1;
  private totalPages: number = 1;
  private objSearchResult: SearchResults;
  private allDataLoaded: boolean = false;
  private allResults: ISearchResult[] = [];

  private PAGE_SIZE: number = 5;
  private columns: IColumn[] = [{
    key: 'column1',
    name: 'Title',
    ariaLabel: 'Column operations for Document Title, Press to sort on Title',
    fieldName: 'title',
    minWidth: 16,
    maxWidth: 160,
    data: 'string',
    onRender: (item: ISearchResult) => (
      <Text>
        <Link href={`${item.link}`} underline={true}>{item.title}</Link>
      </Text>
    )
  },
  // {
  //   key: 'column2',
  //   name: 'Country',
  //   ariaLabel: 'Column operations for country, Press to sort on country',
  //   fieldName: 'Country',
  //   minWidth: 60,
  //   maxWidth: 60,
  //   data: 'string'
  // },
  {
    key: 'column3',
    name: 'Series Number',
    ariaLabel: 'Column operations for country, Press to sort on country',
    fieldName: 'SeriesNumber',
    minWidth: 60,
    maxWidth: 60,
    data: 'string'
  }, {
    key: 'column4',
    name: 'Document Type',
    ariaLabel: 'Column operations for Document Type, Press to sort on Document Type',
    fieldName: 'documentType',
    minWidth: 160,
    maxWidth: 160,
    data: 'string'
  }, {
    key: 'column5',
    name: 'Board Approval Date',
    ariaLabel: 'Column operations for Approved Date, Press to sort on Approved Date',
    fieldName: 'approvedDate',
    minWidth: 160,
    maxWidth: 160,
    onRender: (item: ISearchResult) => (
      item.approvedDate === null ? " " : item.approvedDate.toLocaleDateString()
    )
  }, {
    key: 'column6',
    name: 'Circulation Date',
    ariaLabel: 'Column operations for Circulation Date, Press to sort on Circulation Date',
    fieldName: 'circulationDate',
    minWidth: 160,
    maxWidth: 160,
    onRender: (item: ISearchResult) => (
      item.circulationDate === null ? " " : item.circulationDate.toLocaleDateString()
    )
  }
  ];
  
  private ShowHideChecbox = (event: React.FormEvent<HTMLHeadingElement>) => {
    
    event.currentTarget.nextElementSibling.className = event.currentTarget.nextElementSibling.className == styles.show ? styles.hide : styles.show;
  }

  constructor(props) {
    super(props);
    const result: ISearchResult[] = [];
    const refResult: IRefiners[] = [];

    const temp = this.GetSearchResults("*").then(r => {
      r.forEach(i => {
        result.push(i);
      });
      return result;
    });
    this.state = { searchText: "*", results: result, activePage: 1, refinersResult: refResult, totalRowsCount: 0, seletedprojects: [] , refiners:[], refinementFilters:[]};
  }

  private ExportToExcelOld() {
    const data = [
      {
        data: [
          {
            id: 1,
            level: 0,
            number: '0001',
            name: '0001',
            a: 50,
            b: 45,
            total: 95,
          },
          {
            id: 2,
            parentId: 1,
            level: 1,
            number: '0001-1',
            name: '0001-1',
            a: 20,
            b: 25,
            total: 45,
          },
          {
            id: 3,
            parentId: 2,
            level: 1,
            number: '0001-2',
            name: '0001-2',
            a: 30,
            b: 20,
            total: 50,
          },
          {
            id: 4,
            level: 0,
            number: '0002',
            name: '0002',
            a: 40,
            b: 40,
            total: 80,
          }
        ]
      }
    ];

    const excelExport = new ExcelExport();
    excelExport.downloadExcel(SETTINGS_FOR_EXPORT, data);
  }

  private ExportToExcel() {
    const data = [{ data: this.state.results }];
    const excelExport = new ExcelExport();
    excelExport.downloadExcel(SETTINGS_RESULT_EXPORT, data);
  }

  //GetRefinerString
  public GetRefinersString(selectedOption) {
    let data;
    data = this.state.pnpResults.RawSearchResults.PrimaryQueryResult.RefinementResults.Refiners.filter(k => {
      if (k.Name == selectedOption) {
        return k;
      }
    })[0];
    return data;
  }

  //Get Refiners Value
  public GetRenfinersValue(selectedOption): void {
    console.log(selectedOption);
    this.setState({ chociceOption: [] });
    var arr = [];
    let currentrefiner = this.GetRefinersString(selectedOption);
    currentrefiner.Entries.forEach((entry) => {
      arr.push({ key: entry.RefinementToken, text: entry.RefinementName });
    });
    this.setState({ chociceOption: arr });
  }
  //
  // For pagination
  private GetSearchResultsByPage(pageNumber): Promise<ISearchResult[]> {
    const _resultsRefine: ISearchResult[] = [];
    return new Promise<ISearchResult[]>((resolve, reject) => {
      this.state.pnpResults.getPage(pageNumber, this.PAGE_SIZE)
        .then((results) => {
          results.PrimarySearchResults.forEach((result) => {
            var termString = result["owstaxIdADBBISDocumentType"];
            var termName = termString.split('|')[3].split(';')[0];

            // var termStringCountry= result["owstaxIdADBBISCountry"];
            // var termNameCountry= termStringCountry.split('|')[3].split(';')[0];

            _resultsRefine.push({
              key: result.UniqueId,
              title: result.Title,
              // Country:termNameCountry,
              link: result.Path,
              SeriesNumber: result["ADBBISSeriesNumberOWSTEXT"],
              documentType: termName,
              approvedDate: result["ADBBISApprovalDateOWSDATE"] === null ? null : new Date(result["ADBBISApprovalDateOWSDATE"]),
              circulationDate: result["ADBBISCirculationDateOWSDATE"] === null ? null : new Date(result["ADBBISCirculationDateOWSDATE"])
            });
          });
        })
        .then(
          () => {
            resolve(_resultsRefine);
          }
        )
        .catch(
          () => {
            reject(new Error("Some error"));
          }
        );

    });
  }

  // Get RefineSearchResult
  public RefineSearchResult(refinerName,selectedOption) {
    console.log(refinerName, selectedOption);
    
    this.setState({activePage:0});
    const refinementFilter= this.state.refinementFilters == undefined?[]: this.state.refinementFilters;
    const selectedRefiner=refinementFilter.filter((k)=>{
      if(k.Name == refinerName)
      {
        return k;
      }
    });
    if(selectedRefiner.length == 0 )
    {
      refinementFilter.push({"Name":refinerName, "Filters":[selectedOption]});
    }
    else{
      const filters:string[]= selectedRefiner[0].Filters;
      if(selectedRefiner[0].Filters.indexOf(selectedOption) >= 0)
      {
         selectedRefiner[0].Filters.splice(selectedRefiner[0].Filters.indexOf(selectedOption), 1);
         if(selectedRefiner[0].Filters.length == 0)
         {
          refinementFilter.splice(refinementFilter.indexOf(selectedRefiner[0]) ,1);
         }
      }
      else{
        refinementFilter[refinementFilter.indexOf(selectedRefiner[0])].Filters.push(selectedOption);
      }
    }
    this.setState({refinementFilters:refinementFilter});
    this.refreshSearch(refinementFilter).then(r => { this.setState({ results: r }); });
    
  }

  public refreshSearch(refinementFilter): Promise<ISearchResult[]> {
    
    const refineFilter = refinementFilter.map((k)=>{
     let refinerString : string = "";
      if(k.Filters.length > 1)
      {
        refinerString = "or(" + k.Filters.map(f => {return k.Name + ":" + f}).join(",") + ")"
      }
      else
      {
        refinerString = k.Name + ":" + k.Filters.join(";");
      }
      return refinerString;
    });
    const _results: ISearchResult[] = [];

    return new Promise<ISearchResult[]>((resolve, reject) => {
      sp.search({
        Querytext: this.state.searchText + this.props.queryTemplate,
        SelectProperties: ["Title", "Path", "ADBBISSeriesNumberOWSTEXT", "owstaxIdADBBISDocumentType", "ADBBISApprovalDateOWSDATE", "ADBBISCirculationDateOWSDATE"],
        RowLimit: this.PAGE_SIZE,
        Refiners: "RefinableString00,RefinableString01,RefinableString02,RefinableString03,RefinableString04,RefinableString05,RefinableString06",
        RefinementFilters: refineFilter,
        StartRow: 0
      })
        .then((results) => {
          this.setState({ pnpResults: results, totalRowsCount: results.TotalRows });
          console.log(results.TotalRowsIncludingDuplicates);
          results.PrimarySearchResults.forEach((result) => {
            var termString = result["owstaxIdADBBISDocumentType"];
            var termName = termString.split('|')[3].split(';')[0];

            // var termStringCountry= result["owstaxIdADBBISCountry"];
            // var termNameCountry= termStringCountry.split('|')[3].split(';')[0];

            _results.push({
              key: result.UniqueId,
              title: result.Title,
              link: result.Path,
              // Country: termNameCountry,
              SeriesNumber: result["ADBBISSeriesNumberOWSTEXT"],
              documentType: termName,

              approvedDate: result["ADBBISApprovalDateOWSDATE"] === null ? null : new Date(result["ADBBISApprovalDateOWSDATE"]),
              circulationDate: result["ADBBISCirculationDateOWSDATE"] === null ? null : new Date(result["ADBBISCirculationDateOWSDATE"])
            });
          });

        })
        .then(
          () => {
            resolve(_results);
          }
        )
        .catch(
          (error) => {
            reject(new Error("Some error"));
          }
        );

    });
  }

  public onDropdownMultiChange = async (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption, refinerName: string): Promise<void> => {

    if (item.selected) {
      await arrResult.push(item.key as string);
    }
    else {
      await arrResult.indexOf(item.key) !== -1 && arrResult.splice(arrResult.indexOf(item.key), 1);
    }
    await this.setState({ seletedprojects: arrResult });
    //this.refreshSearch(refinerName + ':' + arrResult.join(',')).then(r => { this.setState({ results: r }); });
    // console.log(this.state.seletedprojects);
  }

  private GetSearchResults(e): Promise<ISearchResult[]> {
    let searchText = "*";
    if (!(e === null || e === undefined || e === ""))
      searchText = e;
    this.setState({ searchText: e });

    const _results: ISearchResult[] = [];
    let _refiners: any[];

    return new Promise<ISearchResult[]>((resolve, reject) => {
      sp.search({
        Querytext: searchText + this.props.queryTemplate,
        SelectProperties: ["Title", "Path", "owstaxIdADBBISCountry", "ADBBISSeriesNumberOWSTEXT", "owstaxIdADBBISDocumentType", "ADBBISApprovalDateOWSDATE", "ADBBISCirculationDateOWSDATE"],
        RowLimit: this.PAGE_SIZE,
        Refiners: "RefinableString00,RefinableString01,RefinableString02,RefinableString03,RefinableString04,RefinableString05,RefinableString06",
        StartRow: 0
      })
        .then((results) => {
          this.setState({ pnpResults: results, totalRowsCount: results.TotalRows });
          console.log(results.TotalRowsIncludingDuplicates);
          results.PrimarySearchResults.forEach((result) => {

            var termString = result["owstaxIdADBBISDocumentType"];
            var termName = termString.split('|')[3].split(';')[0];

            // var termStringCountry= result["owstaxIdADBBISCountry"];
            // var termNameCountry= termStringCountry.split('|')[3].split(';')[0];

            _results.push({
              key: result.UniqueId,
              title: result.Title,
              link: result.Path,
              // Country: result["ADBBISSeriesNumberOWSTEXT"],
              SeriesNumber: result["ADBBISSeriesNumberOWSTEXT"],
              documentType: termName,

              approvedDate: result["ADBBISApprovalDateOWSDATE"] === null ? null : new Date(result["ADBBISApprovalDateOWSDATE"]),
              circulationDate: result["ADBBISCirculationDateOWSDATE"] === null ? null : new Date(result["ADBBISCirculationDateOWSDATE"])
            });
          });
          
          this.CreateRefinerState( results);
        })
        .then(
          () => {
            resolve(_results);
          }
        )
        .catch(
          (error) => {
            reject(new Error("Some error"));
          }
        );

    });
  }

  private CreateRefinerState( results: SearchResults) {
    let _refiners = [];
    if (results.RawSearchResults.PrimaryQueryResult != null) {
      if (results.RawSearchResults.PrimaryQueryResult.RefinementResults != null) {
        if (results.RawSearchResults.PrimaryQueryResult.RefinementResults.Refiners != null) {
          _refiners = results.RawSearchResults.PrimaryQueryResult.RefinementResults.Refiners;
          _refiners.forEach(r => {
            switch (r.Name) {
              case "RefinableString00":
                r.DisplayName = "Country";
                r.Type = "MultiChoice";
                break;
              case "RefinableString01":
                r.DisplayName = "Department";
                r.Type = "MultiChoice";
                break;
              case "RefinableString02":
                r.DisplayName = "Information Classification";
                r.Type = "MultiChoice";
                break;
              case "RefinableString03":
                r.DisplayName = "Series Number";
                r.Type = "MultiChoice";
                break;
              case "RefinableString04":
                r.DisplayName = "Ciruclation Date";
                r.Type = "DateRange";
                break;
              case "RefinableString05":
                r.DisplayName = "Document Type";
                r.Type = "MultiChoice";
                break;
              case "RefinableString06":
                r.DisplayName = "Approval Date";
                r.Type = "DateRange";
                break;
            }
          });
        }
      }
    }
    this.setState({ refiners: _refiners });
    
  }

  public handlePageChange(pageNumber) {
    console.log(`active page is ${pageNumber}`);
    this.GetSearchResultsByPage(pageNumber).then(r => { this.setState({ results: r }); });
    this.setState({ activePage: pageNumber });
  }

  public render(): React.ReactElement<ISpfxSearchProps> {
    const refiners = [...this.state.refiners];
    let refineElement;
    if(refiners != null && refiners.length > 0)
    {
      refineElement = refiners.map(k => {
        return (
          <><div><div className={styles.refinerLable} onClick={this.ShowHideChecbox}>{k.DisplayName}</div>
          <div className={styles.hide} >
            {k.Entries.map(entry => {
              // key: entry.RefinementToken, text: entry.RefinementName
              return(<div><input type="checkbox" value={entry.RefinementToken} onClick={()=>{this.RefineSearchResult(k.Name,entry.RefinementToken)}} /><span>{entry.RefinementName}</span></div>);
            })}
          </div>
          </div></>
        );
      })
    }
   
    const data = this.state.results;
    const searchButtonIcon: IIconProps = { iconName: "Search" };
    const stackTokens: IStackTokens = { childrenGap: 5 };
    return (
      <div className={styles.spfxSearch}>
        <div className={styles.leftDiv}>
          <Stack horizontal={false}>
            <Stack.Item shrink>

              {refineElement}
             
            </Stack.Item>
          </Stack>
        </div>

        <div>
          <Stack horizontal={true} tokens={stackTokens} >
            <Stack.Item>
              <Label htmlFor="searchBox">Search the list:</Label>
            </Stack.Item>
            <Stack.Item grow>
              <SearchBox id="searchBox" placeholder="Please enter the search text"
                aria-label="Search"
                onBlur={(e) => {
                  const value = e.target.value;
                  this.setState({ searchText: value });
                }}
                onSearch={(e) => {
                  // this.GetSearchResults(e).then(r => { this.setState({ results: r }); });
                  this.GetSearchResults(e).then(r => { this.setState({ results: r }); });
                }} >
              </SearchBox>
            </Stack.Item>
            <Stack.Item>
              <DefaultButton aria-label="Search"
                iconProps={searchButtonIcon}
                onClick={(e) => {
                  // this.GetSearchResults(this.state.searchText).then(r => { this.setState({ results: r }); });
                  this.GetSearchResults(this.state.searchText).then(r => { this.setState({ results: r }); });
                }}></DefaultButton>
            </Stack.Item>

          </Stack>

          <Stack>
            <Stack.Item shrink>
              <DefaultButton iconProps={searchButtonIcon} text="Export To Excel"
                onClick={(e) => {
                  this.ExportToExcel();
                }}
                disabled={this.state.results == null || this.state.results.length == 0 ? true : false}>

              </DefaultButton>
            </Stack.Item>
            <Stack.Item>
              <DetailsList
                items={this.state.results}
                columns={this.columns}
                selectionMode={SelectionMode.none}
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                selectionPreservedOnEmptyClick={true}
              />
            </Stack.Item>
            <Stack.Item>
              <Pagination
                activePage={this.state.activePage}
                itemsCountPerPage={this.PAGE_SIZE}
                totalItemsCount={this.state.totalRowsCount}
                pageRangeDisplayed={5}
                itemClass="page-item"
                linkClass="page-link"
                onChange={this.handlePageChange.bind(this)}
              />
            </Stack.Item>
          </Stack>

        </div >
      </div>
    );
  }
}
