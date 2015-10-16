/*
 *  Power BI Visualizations
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved. 
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *   
 *  The above copyright notice and this permission notice shall be included in 
 *  all copies or substantial portions of the Software.
 *   
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */

/// <reference path="../../_references.ts"/>

module powerbi.visuals.samples {
    //import SelectionManager = utility.SelectionManager;
    import ClassAndSelector = jsCommon.CssConstants.ClassAndSelector;

    import PixelConverter = jsCommon.PixelConverter;

    export interface SlicerXConstructorOptions {
        behavior?: SlicerXWebBehavior;
        svg?: D3.Selection;
        margin?: IMargin;
    }

    export interface SlicerXData {
        categorySourceName: string;
        formatString: string;
        slicerDataPoints: SlicerXDataPoint[];
        slicerSettings: SlicerXSettings;
        hasSelectionOverride?: boolean;
    }

    export interface SlicerXDataPoint extends SelectableDataPoint {
        category?: string;
        value: string;
        mouseOver: boolean;
        mouseOut: boolean;
        isSelectAllDataPoint?: boolean;
        imageURL?: string;
    }

    export interface SlicerXSettings {
        general: {
            horizontal: boolean;
            columns: number;
            //  sortorder: string;
            multiselect: boolean;
            showdisabled: string;
        };
        margin: IMargin;
        header: {
            borderBottomWidth: number;
            show: boolean;
            outline: string;
            fontColor: string;
            background: string;
            textSize: number;
            outlineColor: string;
            outlineWeight: number;
            title: string;
        };
        headerText: {
            marginLeft: number;
            marginTop: number;
        };
        slicerText: {
            textSize: number;
            height: number;
            width: number;
            fontColor: string;
            hoverColor: string;
            selectedColor: string;
            unselectedColor: string;
            disabledColor: string;
            marginLeft: number;
            outline: string;
            background: string;
            outlineColor: string;
            outlineWeight: number;
        };
        slicerItemContainer: {
            marginTop: number;
            marginLeft: number;
        };
        images: {
            imageSplit: number;
            stretchImage: boolean;
            bottomImage: boolean;
        };
    }

    export class SlicerX implements IVisual {
        private element: JQuery;
        private currentViewport: IViewport;
        private dataView: DataView;
        private slicerHeader: D3.Selection;
        private slicerBody: D3.Selection;
        private tableView: ITableView;
        private slicerData: SlicerXData;
        private settings: SlicerXSettings;
        private interactivityService: IInteractivityService;
        private behavior: SlicerXWebBehavior;
        private hostServices: IVisualHostServices;
        //private static clearTextKey = 'Slicer_Clear';
        //private static selectAllTextKey = 'Slicer_SelectAll';
        private waitingForData: boolean;
        private textProperties: TextProperties = {
            'fontFamily': 'wf_segoe-ui_normal, helvetica, arial, sans-serif',
            'fontSize': '14px',
        };

        private static Container: ClassAndSelector = {
            class: 'slicerX',
            selector: '.slicerX'
        };
        private static Header: ClassAndSelector = {
            class: 'slicerHeader',
            selector: '.slicerHeader'
        };
        private static HeaderText: ClassAndSelector = {
            class: 'headerText',
            selector: '.headerText'
        };
        private static Body: ClassAndSelector = {
            class: 'slicerBody',
            selector: '.slicerBody'
        };
        private static ItemContainer: ClassAndSelector = {
            class: 'slicerItemContainer',
            selector: '.slicerItemContainer'
        };
        private static LabelText: ClassAndSelector = {
            class: 'slicerText',
            selector: '.slicerText'
        };
        private static Input: ClassAndSelector = {
            class: 'slicerCheckbox',
            selector: '.slicerCheckbox'
        };
        private static Clear: ClassAndSelector = {
            class: 'clear',
            selector: '.clear'
        };

        public static DefaultStyleProperties(): SlicerXSettings {
            return {
                general: {
                    horizontal: false,
                    columns: 0,
                    //  sortorder: 'ASC',
                    multiselect: true,
                    showdisabled: 'In-Place'
                },
                margin: {
                    top: 50,
                    bottom: 50,
                    right: 50,
                    left: 50
                },
                header: {
                    borderBottomWidth: 1,
                    show: true,
                    outline: 'BottomOnly',
                    fontColor: '#000000',
                    background: '#ffffff',
                    textSize: 10,
                    outlineColor: '#000000',
                    outlineWeight: 1,
                    title: '',
                },
                headerText: {
                    marginLeft: 8,
                    marginTop: 0
                },
                slicerText: {
                    textSize: 10,
                    height: 0,
                    width: 0,
                    fontColor: '#666666',
                    hoverColor: '#212121',
                    selectedColor: '#BDD7EE',
                    unselectedColor: '#ffffff',
                    disabledColor: 'grey',
                    marginLeft: 8,
                    outline: 'Frame',
                    background: '#ffffff',
                    outlineColor: '#000000',
                    outlineWeight: 1,

                },
                slicerItemContainer: {
                    // The margin is assigned in the less file. This is needed for the height calculations.
                    marginTop: 5,
                    marginLeft: 0,
                },
                images: {
                    imageSplit: 50,
                    stretchImage: false,
                    bottomImage: false
                }
            };
        }

        private svg: D3.Selection;

        constructor(options?: SlicerXConstructorOptions) {
            if (options) {
                if (options.svg) {
                    this.svg = options.svg;
                }
                if (options.behavior) {
                    this.behavior = options.behavior;
                }
            }
        }

        public static converter(dataView: DataView, localizedSelectAllText: string, interactivityService: IInteractivityService): SlicerXData {
            let slicerData: SlicerXData;
            if (!dataView) {
                return;
            }

            let dataViewCategorical = dataView.categorical;
            if (dataViewCategorical == null || dataViewCategorical.categories == null || dataViewCategorical.categories.length === 0)
                return;

            let isInvertedSelectionMode = undefined;
            let objects = dataView.metadata ? <any>dataView.metadata.objects : undefined;
            let categories = dataViewCategorical.categories[0];
            let categoryValues = dataViewCategorical.values[0];

            let numberOfScopeIds: number;
            if (objects && objects.general && objects.general.filter) {
                let identityFields = categories.identityFields;
                if (!identityFields)
                    return;
                let filter = <powerbi.data.SemanticFilter>objects.general.filter;
                let scopeIds = powerbi.data.SQExprConverter.asScopeIdsContainer(filter, identityFields);
                if (scopeIds) {
                    isInvertedSelectionMode = scopeIds.isNot;
                    numberOfScopeIds = scopeIds.scopeIds ? scopeIds.scopeIds.length : 0;
                }
                else {
                    isInvertedSelectionMode = false;
                }
            }

            if (interactivityService) {
                if (isInvertedSelectionMode === undefined) {
                    // The selection state is read from the Interactivity service in case of SelectAll or Clear when query doesn't update the visual
                    isInvertedSelectionMode = interactivityService.isSelectionModeInverted();
                }
                else {
                    interactivityService.setSelectionModeInverted(isInvertedSelectionMode);
                }
            }

            let categoryValuesLen = categories && categories.values ? categories.values.length : 0;
            let slicerDataPoints: SlicerXDataPoint[] = [];

            //slicerDataPoints.push({
            //    value: localizedSelectAllText,
            //    mouseOver: false,
            //    mouseOut: true,
            //    identity: SelectionId.createWithMeasure(localizedSelectAllText),
            //    selected: !!isInvertedSelectionMode,
            //    isSelectAllDataPoint: true
            //});                    
                                     
            // Pass over the values to see if there's a positive or negative selection
            let hasSelection: boolean = undefined;

            for (let idx = 0; idx < categoryValuesLen; idx++) {
                let selected = isCategoryColumnSelected(slicerXProps.selectedPropertyIdentifier, categories, idx);
                if (selected != null) {
                    hasSelection = selected;
                    break;
                }
            }

            let numberOfCategoriesSelectedInData = 0;
            for (let idx = 0; idx < categoryValuesLen; idx++) {
                let categoryIdentity = categories.identity ? categories.identity[idx] : null;
                let categoryIsSelected = isCategoryColumnSelected(slicerXProps.selectedPropertyIdentifier, categories, idx);

                if (hasSelection != null) {
                    // If the visual is in InvertedSelectionMode, all the categories should be selected by default unless they are not selected
                    // If the visual is not in InvertedSelectionMode, we set all the categories to be false except the selected category                         
                    if (isInvertedSelectionMode) {
                        if (categories.objects == null)
                            categoryIsSelected = undefined;

                        if (categoryIsSelected != null) {
                            categoryIsSelected = hasSelection;
                        }
                        else if (categoryIsSelected == null)
                            categoryIsSelected = !hasSelection;
                    }
                    else {
                        if (categoryIsSelected == null) {
                            categoryIsSelected = !hasSelection;
                        }
                    }
                }

                if (categoryIsSelected)
                    numberOfCategoriesSelectedInData++;

                let dataPoint: SlicerXDataPoint = {
                    value: categories.values[idx],
                    mouseOver: false,
                    mouseOut: true,
                    identity: SelectionId.createWithId(categoryIdentity),
                    selected: categoryIsSelected
                };
                /*
                if (categoryValues && categoryValues[idx] && categoryValues[idx].source) {
                    dataPoint.imageURL = converterHelper.getFormattedLegendLabel(categoryValues[idx].source, dataViewCategorical.values, slicerXProps.formatString);
                }*/
                dataPoint.imageURL = categoryValues.values[idx];
                slicerDataPoints.push(dataPoint);
            }

            let defaultSettings = this.DefaultStyleProperties();
            objects = dataView.metadata.objects;
            if (objects) {
                defaultSettings.general.horizontal = DataViewObjects.getValue<boolean>(objects, slicerXProps.general.horizontal, defaultSettings.general.horizontal);
                defaultSettings.general.columns = DataViewObjects.getValue<number>(objects, slicerXProps.general.columns, defaultSettings.general.columns);
                //   defaultSettings.general.sortorder = DataViewObjects.getValue<string>(objects, slicerXProps.general.sortorder, defaultSettings.general.sortorder);
                defaultSettings.general.multiselect = DataViewObjects.getValue<boolean>(objects, slicerXProps.general.multiselect, defaultSettings.general.multiselect);
                defaultSettings.general.showdisabled = DataViewObjects.getValue<string>(objects, slicerXProps.general.showdisabled, defaultSettings.general.showdisabled);

                defaultSettings.header.show = DataViewObjects.getValue<boolean>(objects, slicerXProps.header.show, defaultSettings.header.show);
                defaultSettings.header.title = DataViewObjects.getValue<string>(objects, slicerXProps.header.title, defaultSettings.header.title);
                defaultSettings.header.fontColor = DataViewObjects.getFillColor(objects, slicerXProps.header.fontColor, defaultSettings.header.fontColor);
                defaultSettings.header.background = DataViewObjects.getFillColor(objects, slicerXProps.header.background, defaultSettings.header.background);
                defaultSettings.header.textSize = DataViewObjects.getValue<number>(objects, slicerXProps.header.textSize, defaultSettings.header.textSize);
                defaultSettings.header.outline = DataViewObjects.getValue<string>(objects, slicerXProps.header.outline, defaultSettings.header.outline);
                defaultSettings.header.outlineColor = DataViewObjects.getFillColor(objects, slicerXProps.header.outlineColor, defaultSettings.header.outlineColor);
                defaultSettings.header.outlineWeight = DataViewObjects.getValue<number>(objects, slicerXProps.header.outlineWeight, defaultSettings.header.outlineWeight);

                defaultSettings.slicerText.textSize = DataViewObjects.getValue<number>(objects, slicerXProps.rows.textSize, defaultSettings.slicerText.textSize);
                defaultSettings.slicerText.height = DataViewObjects.getValue<number>(objects, slicerXProps.rows.height, defaultSettings.slicerText.height);
                defaultSettings.slicerText.width = DataViewObjects.getValue<number>(objects, slicerXProps.rows.width, defaultSettings.slicerText.width);
                defaultSettings.slicerText.selectedColor = DataViewObjects.getFillColor(objects, slicerXProps.rows.selectedColor, defaultSettings.slicerText.selectedColor);
                defaultSettings.slicerText.unselectedColor = DataViewObjects.getFillColor(objects, slicerXProps.rows.unselectedColor, defaultSettings.slicerText.unselectedColor);
                defaultSettings.slicerText.disabledColor = DataViewObjects.getFillColor(objects, slicerXProps.rows.disabledColor, defaultSettings.slicerText.disabledColor);
                defaultSettings.slicerText.background = DataViewObjects.getFillColor(objects, slicerXProps.rows.background, defaultSettings.slicerText.background);
                defaultSettings.slicerText.fontColor = DataViewObjects.getFillColor(objects, slicerXProps.rows.fontColor, defaultSettings.slicerText.fontColor);
                defaultSettings.slicerText.outline = DataViewObjects.getValue<string>(objects, slicerXProps.rows.outline, defaultSettings.slicerText.outline);
                defaultSettings.slicerText.outlineColor = DataViewObjects.getFillColor(objects, slicerXProps.rows.outlineColor, defaultSettings.slicerText.outlineColor);
                defaultSettings.slicerText.outlineWeight = DataViewObjects.getValue<number>(objects, slicerXProps.rows.outlineWeight, defaultSettings.slicerText.outlineWeight);

                defaultSettings.images.imageSplit = DataViewObjects.getValue<number>(objects, slicerXProps.images.imageSplit, defaultSettings.images.imageSplit);
                defaultSettings.images.stretchImage = DataViewObjects.getValue<boolean>(objects, slicerXProps.images.stretchImage, defaultSettings.images.stretchImage);
                defaultSettings.images.bottomImage = DataViewObjects.getValue<boolean>(objects, slicerXProps.images.bottomImage, defaultSettings.images.bottomImage);
            }

            slicerData = {
                categorySourceName: categories.source.displayName,
                formatString: valueFormatter.getFormatString(categories.source, slicerXProps.formatString),
                slicerSettings: defaultSettings,
                slicerDataPoints: slicerDataPoints,
            };

            // Override hasSelection if a objects contained more scopeIds than selections we found in the data
            if (numberOfScopeIds != null && numberOfScopeIds > numberOfCategoriesSelectedInData) {
                slicerData.hasSelectionOverride = true;
            }

            return slicerData;
        }

        public init(options: VisualInitOptions): void {
            this.element = options.element;
            this.currentViewport = options.viewport;
            if (this.behavior) {
                this.interactivityService = createInteractivityService(options.host);
            }
            this.hostServices = options.host;
            this.settings = SlicerX.DefaultStyleProperties();

            this.initContainer();
        }

        public onDataChanged(options: VisualDataChangedOptions): void {
            let dataViews = options.dataViews;
            debug.assertValue(dataViews, 'dataViews');

            let existingDataView = this.dataView;
            if (dataViews && dataViews.length > 0) {
                this.dataView = dataViews[0];
            }

            let resetScrollbarPosition = false;
            // Null check is needed here. If we don't check for null, selecting a value on loadMore event will evaluate the below condition to true and resets the scrollbar
            if (options.operationKind !== undefined) {
                resetScrollbarPosition = options.operationKind !== VisualDataChangeOperationKind.Append
                && !DataViewAnalysis.hasSameCategoryIdentity(existingDataView, this.dataView);
            }

            this.updateInternal(resetScrollbarPosition);
            this.waitingForData = false;
        }

        public onResizing(finalViewport: IViewport): void {
            this.currentViewport = finalViewport;
            this.updateInternal(false /* resetScrollbarPosition */);
        }

        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] {
            let data = this.slicerData;
            if (!data)
                return;

            let objectName = options.objectName;
            switch (objectName) {
                case 'rows':
                    return this.enumerateRows(data);
                case 'header':
                    return this.enumerateHeader(data);
                case 'general':
                    return this.enumerateGeneral(data);
                case 'images':
                    return this.enumerateImages(data);
            }
        }

        private enumerateHeader(data: SlicerXData): VisualObjectInstance[] {
            let slicerSettings = this.settings;
            //let fontColor = data !== undefined && data.slicerSettings !== undefined && data.slicerSettings.header && data.slicerSettings.header.fontColor ?
            //   data.slicerSettings.header.fontColor : slicerSettings.header.fontColor;
            // let background = data !== undefined && data.slicerSettings !== undefined && data.slicerSettings.header && data.slicerSettings.header.background ?
            //    data.slicerSettings.header.background : slicerSettings.header.background;
            return [{
                selector: null,
                objectName: 'header',
                properties: {
                    show: slicerSettings.header.show,
                    title: slicerSettings.header.title,
                    fontColor: slicerSettings.header.fontColor,
                    background: slicerSettings.header.background,
                    textSize: slicerSettings.header.textSize,
                    outline: slicerSettings.header.outline,
                    outlineColor: slicerSettings.header.outlineColor,
                    outlineWeight: slicerSettings.header.outlineWeight

                }
            }];
        }

        private enumerateRows(data: SlicerXData): VisualObjectInstance[] {
            let slicerSettings = this.settings;
            //let fontColor = data !== undefined && data.slicerSettings !== undefined && data.slicerSettings.slicerText && data.slicerSettings.slicerText.color ?
            //    data.slicerSettings.slicerText.color : slicerSettings.slicerText.color;
            //let background = data !== undefined && data.slicerSettings !== undefined && data.slicerSettings.slicerText && data.slicerSettings.slicerText.background ?
            //    data.slicerSettings.slicerText.background : slicerSettings.slicerText.background;
            return [{
                selector: null,
                objectName: 'rows',
                properties: {
                    textSize: slicerSettings.slicerText.textSize,
                    height: slicerSettings.slicerText.height,
                    width: slicerSettings.slicerText.width,
                    background: slicerSettings.slicerText.background,
                    selectedColor: slicerSettings.slicerText.selectedColor,
                    unselectedColor: slicerSettings.slicerText.unselectedColor,
                    disabledColor: slicerSettings.slicerText.disabledColor,
                    outline: slicerSettings.slicerText.outline,
                    outlineColor: slicerSettings.slicerText.outlineColor,
                    outlineWeight: slicerSettings.slicerText.outlineWeight,
                    fontColor: slicerSettings.slicerText.fontColor,
                }
            }];
        }

        private enumerateGeneral(data: SlicerXData): VisualObjectInstance[] {
            let slicerSettings = this.settings;
            //let outlineColor = data !== undefined && data.slicerSettings !== undefined && data.slicerSettings.general && data.slicerSettings.general.outlineColor ?
            //    data.slicerSettings.general.outlineColor : slicerSettings.general.outlineColor;
            //let outlineWeight = data !== undefined && data.slicerSettings !== undefined && data.slicerSettings.general && data.slicerSettings.general.outlineWeight ?
            //    data.slicerSettings.general.outlineWeight : slicerSettings.general.outlineWeight;

            return [{
                selector: null,
                objectName: 'general',
                properties: {
                    horizontal: slicerSettings.general.horizontal,
                    columns: slicerSettings.general.columns,
                    //   sortorder: slicerSettings.general.sortorder,
                    multiselect: slicerSettings.general.multiselect,
                    showdisabled: slicerSettings.general.showdisabled,
                }
            }];
        }

        private enumerateImages(data: SlicerXData): VisualObjectInstance[] {
            let slicerSettings = this.settings;
            //let outlineColor = data !== undefined && data.slicerSettings !== undefined && data.slicerSettings.general && data.slicerSettings.general.outlineColor ?
            //    data.slicerSettings.general.outlineColor : slicerSettings.general.outlineColor;
            //let outlineWeight = data !== undefined && data.slicerSettings !== undefined && data.slicerSettings.general && data.slicerSettings.general.outlineWeight ?
            //    data.slicerSettings.general.outlineWeight : slicerSettings.general.outlineWeight;

            return [{
                selector: null,
                objectName: 'images',
                properties: {
                    imageSplit: slicerSettings.images.imageSplit,
                    stretchImage: slicerSettings.images.stretchImage,
                    bottomImage: slicerSettings.images.bottomImage,
                }
            }];
        }

        private updateInternal(resetScrollbarPosition: boolean) {
            this.updateSlicerBodyDimensions();

            let localizedSelectAllText = 'Select All';
            let data = SlicerX.converter(this.dataView, localizedSelectAllText, this.interactivityService);
            if (!data) {
                this.tableView.empty();
                return;
            }

            data.slicerSettings.header.outlineWeight = data.slicerSettings.header.outlineWeight < 0 ? 0 : data.slicerSettings.header.outlineWeight;
            this.slicerData = data;
            this.settings = this.slicerData.slicerSettings;
            //  this.tableView.empty();
            this.tableView
                .viewport(this.getSlicerBodyViewport(this.currentViewport))
                .rowHeight(this.settings.slicerText.height)
                .columnWidth(this.settings.slicerText.width)
                .rows(this.settings.general.horizontal ? 1 : 0)
                .columns(this.settings.general.columns)
                .data(
                    data.slicerDataPoints,
                    (d: SlicerXDataPoint) => $.inArray(d, data.slicerDataPoints),
                    resetScrollbarPosition
                );
        }

        private initContainer() {
            let settings = this.settings;
            let slicerBodyViewport = this.getSlicerBodyViewport(this.currentViewport);
            let slicerContainer: D3.Selection = d3.select(this.element.get(0)).classed(SlicerX.Container.class, true);

            this.slicerHeader = slicerContainer.append('div').classed(SlicerX.Header.class, true);

            this.slicerHeader.append('span')
                .classed(SlicerX.Clear.class, true)
                .attr('title', 'Clear');

            this.slicerHeader.append('div').classed(SlicerX.HeaderText.class, true)
                .style({
                    'margin-left': PixelConverter.toString(settings.headerText.marginLeft),
                    'margin-top': PixelConverter.toString(settings.headerText.marginTop),
                    'border-style': this.getBorderStyle(settings.header.outline),
                    'border-color': settings.header.outlineColor,
                    'border-width': this.getBorderWidth(settings.header.outline, settings.header.outlineWeight),
                    'font-size': PixelConverter.fromPoint(settings.header.textSize),
                });

            this.slicerBody = slicerContainer.append('div').classed(SlicerX.Body.class, true).classed('slicerBody-horizontal', settings.general.horizontal)
                .style({
                    'height': PixelConverter.toString(slicerBodyViewport.height),
                    'width': PixelConverter.toString(slicerBodyViewport.width),
                });
            //this.slicerBody.append('div').classed('slicer-wrapper', true)
            //    .style({
            //        'width': 'auto',
            //        'white-space': 'nowrap'
            //    });
            let rowEnter = (rowSelection: D3.Selection) => {
                let settings = this.settings;
//                let labelWidth = PixelConverter.toString(this.currentViewport.width - (settings.slicerItemContainer.marginLeft + settings.slicerText.marginLeft + settings.header.outlineWeight * 2));
                //  rowSelection.classed('row-horizontal', settings.general.horizontal);
                let listItemElement = rowSelection.append('li')
                    .classed(SlicerX.ItemContainer.class, true)
                    .style({
                        'margin-left': PixelConverter.toString(settings.slicerItemContainer.marginLeft),
                    });

                //let labelElement = listItemElement.append('label')
                //    .classed(SlicerX.Input.class, true);

                //labelElement.append('input')
                //    .attr('type', 'checkbox');
                listItemElement.append('div').style('display', 'none')
                    .classed('slicer-img-wrapper', true)
                    .append('img')
                    .classed('slicer-img', true);

                listItemElement.append('div')
                    .classed('slicer-text-wrapper', true)
                    .append('span')
                    .classed(SlicerX.LabelText.class, true)
                    .style({
                        //  'width': settings.gener7al.horizontal === true ? labelWidth : 'auto',
                        'font-size': PixelConverter.fromPoint(settings.slicerText.textSize),
                    });

            };

            let rowUpdate = (rowSelection: D3.Selection) => {
                let settings = this.settings;
                let data = this.slicerData;
                if (data && settings) {

                    if (settings.header.show) {
                        this.slicerHeader.style('display', 'block');
                        this.slicerHeader.select(SlicerX.HeaderText.selector)
                            .text(settings.header.title.trim() !== "" ? settings.header.title.trim() : this.slicerData.categorySourceName)
                            .style({
                                'border-style': this.getBorderStyle(settings.header.outline),
                                'border-color': settings.header.outlineColor,
                                'border-width': this.getBorderWidth(settings.header.outline, settings.header.outlineWeight),
                                'color': settings.header.fontColor,
                                'background-color': settings.header.background,
                                'font-size': PixelConverter.fromPoint(settings.header.textSize),
                            });
                    }
                    else {
                        this.slicerHeader.style('display', 'none');
                    }

                    let slicerText = rowSelection.selectAll(SlicerX.LabelText.selector);

                    let formatString = data.formatString;
                    slicerText.text((d: SlicerXDataPoint) => valueFormatter.format(d.value, formatString));
                    let slicerImg = rowSelection.selectAll('.slicer-img-wrapper').style('display', (d: SlicerXDataPoint) => d.imageURL ? 'inline-block' : 'none')
                        .style('width', settings.images.imageSplit + '%');
                    slicerImg.selectAll('img').attr('src', (d: SlicerXDataPoint) => d.imageURL)
                        .style({
                            'width': settings.images.stretchImage ? '100%' : 'auto',
                            'height': settings.images.stretchImage ? '100%' : 'auto',
                        });;
                    rowSelection.selectAll('.slicer-text-wrapper').style('width', (d: SlicerXDataPoint) => d.imageURL ? (100 - settings.images.imageSplit) + '%' : '100%');
                    rowSelection.style({
                        'color': settings.slicerText.fontColor,
                        'border-style': this.getBorderStyle(settings.slicerText.outline),
                        'border-color': settings.slicerText.outlineColor,
                        'border-width': this.getBorderWidth(settings.slicerText.outline, settings.slicerText.outlineWeight),
                        'font-size': PixelConverter.fromPoint(settings.slicerText.textSize),
                    });

                    if (this.interactivityService && this.slicerBody) {
                        let slicerBody = this.slicerBody.attr('width', this.currentViewport.width);
                        let slicerItemContainers = slicerBody.selectAll(SlicerX.ItemContainer.selector);
                        let slicerItemLabels = slicerBody.selectAll(SlicerX.LabelText.selector);
                        let slicerItemInputs = slicerBody.selectAll(SlicerX.Input.selector);
                        let slicerClear = this.slicerHeader.select(SlicerX.Clear.selector);

                        let behaviorOptions: SlicerXBehaviorOptions = {
                            dataPoints: data.slicerDataPoints,
                            slicerItemContainers: slicerItemContainers,
                            slicerItemLabels: slicerItemLabels,
                            slicerItemInputs: slicerItemInputs,
                            slicerClear: slicerClear,
                            interactivityService: this.interactivityService,
                            slicerSettings: data.slicerSettings,
                        };

                        this.interactivityService.bind(data.slicerDataPoints, this.behavior, behaviorOptions, { overrideSelectionFromData: true, hasSelectionOverride: data.hasSelectionOverride });
                        this.behavior.styleSlicerInputs(rowSelection.select(SlicerX.ItemContainer.selector), this.interactivityService.hasSelection());
                    }
                    else {
                        this.behavior.styleSlicerInputs(rowSelection.select(SlicerX.ItemContainer.selector), false);
                    }
                }
            };

            let rowExit = (rowSelection: D3.Selection) => {
                rowSelection.remove();
            };

            let tableViewOptions: TableViewViewOptions = {
                rowHeight: this.getRowHeight(),
                columnWidth: this.settings.slicerText.width,
                rows: this.settings.general.horizontal ? 1 : 0,
                columns: this.settings.general.columns,
                enter: rowEnter,
                exit: rowExit,
                update: rowUpdate,
                loadMoreData: () => this.onLoadMoreData(),
                scrollEnabled: true,
                viewport: this.getSlicerBodyViewport(this.currentViewport),
                baseContainer: this.slicerBody,
            };

            this.tableView = TableViewFactory.createTableView(tableViewOptions);
        }

        private onLoadMoreData(): void {
            if (!this.waitingForData && this.dataView.metadata && this.dataView.metadata.segment) {
                this.hostServices.loadMoreData();
                this.waitingForData = true;
            }
        }

        private getSlicerBodyViewport(currentViewport: IViewport): IViewport {
            let settings = this.settings;
            let headerHeight = (settings.header.show) ? this.getHeaderHeight() : 0;
            let slicerBodyHeight = currentViewport.height - (headerHeight + settings.header.borderBottomWidth);
            return {
                height: slicerBodyHeight,
                width: currentViewport.width
            };
        }

        private updateSlicerBodyDimensions(): void {
            let slicerViewport = this.getSlicerBodyViewport(this.currentViewport);
            this.slicerBody
                .style({
                    'height': PixelConverter.toString(slicerViewport.height),
                    'width': PixelConverter.toString(slicerViewport.width),
                });
        }

        private getTextProperties(textSize: number): TextProperties {
            this.textProperties.fontSize = PixelConverter.fromPoint(textSize);
            return this.textProperties;
        }

        private getHeaderHeight(): number {
            return TextMeasurementService.estimateSvgTextHeight(
                this.getTextProperties(this.settings.header.textSize)
            );
        }

        private getRowHeight(): number {
            let textSettings = this.settings.slicerText;
            return textSettings.height !== 0 ? textSettings.height : TextMeasurementService.estimateSvgTextHeight(
                this.getTextProperties(textSettings.textSize)
            );
        }

        private getBorderStyle(outlineElement: string): string {
            return outlineElement === '0px' ? 'none' : 'solid';
        }

        private getBorderWidth(outlineElement: string, outlineWeight: number): string {
            switch (outlineElement) {
                case 'None':
                    return '0px';
                case 'BottomOnly':
                    return '0px 0px ' + outlineWeight + 'px 0px';
                case 'TopOnly':
                    return outlineWeight + 'px 0px 0px 0px';
                case 'TopBottom':
                    return outlineWeight + 'px 0px ' + outlineWeight + 'px 0px';
                case 'LeftRight':
                    return '0px ' + outlineWeight + 'px 0px ' + outlineWeight + 'px';
                case 'Frame':
                    return outlineWeight + 'px';
                default:
                    return outlineElement.replace("1", outlineWeight.toString());

            }
        }

        //private static getMetadataFill(dataView: DataView, field: string, property: string, defaultValue: string): Fill {
        //    if (dataView) {
        //        let metadata = dataView.metadata.objects;
        //        if (metadata) {
        //            let val = metadata[field];
        //            if (val && val.hasOwnProperty(property)) {
        //                let fill = <Fill>val[property];
        //                if (fill)
        //                    return fill;
        //            }
        //        }
        //    }
        //    return { solid: { color: defaultValue } };
        //}

        //private static getMetadataText(dataView: DataView, field: string, property: string, defaultValue: string = ''): string {
        //    if (dataView) {
        //        let metadata = dataView.metadata.objects;
        //        if (metadata) {
        //            let val = metadata[field];
        //            if (val && val.hasOwnProperty(property) && val[property] !== '') {
        //                let text = <string>val[property];
        //                if (text)
        //                    return text;
        //            }
        //        }
        //    }
        //    return defaultValue;
        //}

        //private static getMetadataBool(dataView: DataView, field: string, property: string, defaultValue: boolean = true): boolean {
        //    if (dataView) {
        //        let metadata = dataView.metadata.objects;
        //        if (metadata) {
        //            let val = metadata[field];
        //            if (val && val.hasOwnProperty(property))
        //                return <boolean>val[property];
        //        }
        //    }
        //    return defaultValue;
        //}

        //private static getMetadataNumber(dataView: DataView, field: string, property: string, defaultValue: number = -Infinity): number {
        //    if (dataView) {
        //        let metadata = dataView.metadata.objects;
        //        if (metadata) {
        //            let val = metadata[field];
        //            if (val && val.hasOwnProperty(property) && val[property] !== '')
        //                return <number>val[property];
        //        }
        //    }
        //    return defaultValue;
        //}

    }
}
