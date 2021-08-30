import { IColumn } from '@fluentui/react/lib/DetailsList';

export interface IDetailsListSearchResultState {
    columns: IColumn[];
    items: ISearchResult[];
    selectionDetails: string;
    isModalSelection: boolean;
    isCompactMode: boolean;
    announcedMessage?: string;
   
}

export interface ISearchResult {

    key: string;
    title: string;
   // Country: string;
    SeriesNumber: string;
    documentType: string;
    approvedDate: Date;
    circulationDate: Date;
    link: string;
}

export interface IRefiners {
    key: string;
    title: string;
    country : string;
    department : string;
    InfoClassification: string;
    SeriesNumber: string;
    CirculationDate: Date;
    ApprovalDate: Date;
    link:string;
    
}
export interface IRefinerFilter {
    Name: string;
    Filters: string[];    
}
