import { alignment, defaultDataType } from 'export-xlsx';

// Export settings
export const SETTINGS_FOR_EXPORT = {
    // Table settings
    fileName: 'example',
    workSheets: [
        {
            sheetName: 'example',
            startingRowNumber: 2,
            gapBetweenTwoTables: 2,
            tableSettings: {
                data: {
                    importable: true,
                    tableTitle: 'Score',
                    notification: 'Notify: only yellow background cell could edit!',
                    headerGroups: [
                        {
                            name: 'Score',
                            key: 'score',
                        },
                    ],
                    headerDefinition: [
                        {
                            name: 'Id',
                            key: 'id',
                            width: 25,
                            hierarchy: true,
                            checkable: true,
                        },
                        {
                            name: 'Number',
                            key: 'number',
                            width: 18,
                            checkable: true,
                            style: { alignment: alignment.middleCenter },
                        },
                        {
                            name: 'Name',
                            key: 'name',
                            width: 18,
                            style: { alignment: alignment.middleCenter },
                        },
                        {
                            name: 'A',
                            key: 'a',
                            width: 18,
                            groupKey: 'score',
                            dataType: defaultDataType.number,
                            selfSum: true,
                            editable: true,
                        },
                        {
                            name: 'B',
                            key: 'b',
                            width: 18,
                            groupKey: 'score',
                            dataType: defaultDataType.number,
                            selfSum: true,
                            editable: true,
                        },
                        {
                            name: 'Total',
                            key: 'total',
                            width: 18,
                            dataType: defaultDataType.number,
                            selfSum: true,
                            rowFormula: '{a}+{b}',
                        },
                    ],
                },
            },
        },
    ],
};

export const SETTINGS_RESULT_EXPORT = {
    fileName: "SearchResult",
    workSheets: [
        {
            sheetName: "Results",
            startingRowNumber: 1,
            gapBetweenTwoTables: 1,
            tableSettings: {
                data: {
                    importable: true,
                    headerDefinition: [
                        {
                            name: 'Title',
                            key: 'title',
                            width: 25,
                            style: {
                                alignment: alignment.middleLeft
                            },
                        },
                        {
                            name: 'Series Number',
                            key: 'SeriesNumber',
                            width: 25,
                            style: {
                                alignment: alignment.middleLeft
                            },
                        },
                        {
                            name: 'Document Type',
                            key: 'documentType',
                            width: 25,
                            style: {
                                alignment: alignment.middleLeft
                            },
                        },
                        {
                            name: 'Approved Date',
                            key: 'approvedDate',
                            width: 25,
                            dataType: defaultDataType.date,
                            style: {
                                alignment: alignment.middleLeft
                            },
                        },
                        {
                            name: 'Circulation Date',
                            key: 'circulationDate',
                            width: 25,
                            dataType: defaultDataType.date,
                            style: {
                                alignment: alignment.middleLeft
                            },
                        }
                    ]
                }
            }
        }
    ]
};