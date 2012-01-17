Ext.ns('Ext.ux.grid');
/**
 * @author Shea Frederick - http://www.vinylfox.com
 * @class Ext.ux.grid.XmlExport
 * @extends Ext.AbstractPlugin
 * @contributor Adapted from code written by Nigel White (aka Animal)
 *
 * <p>A plugin that can create an XML file (compatible with Excel) for download
 * based on selections in the grid, or just the entire grids contents.</p>
 * <pre><code>
 {
     xtype: 'grid',
     ...,
     plugins: [{ptype:'xmlexport'}],
     ...
 }
 * </code></pre>
 * 
 * @alias plugin.xmlexport
 * @ptype xmlexport
 */
Ext.define('Ext.ux.grid.XmlExport', {
    extend: 'Ext.util.Observable',
    alias: 'plugin.xmlexport',
    localeExportAll: 'Export All',
    localeExportSelection: 'Export Selection',
    /**
     * @cfg showExportSelection {boolean} set true to show the export selection button (defaults to true). 
     */
    showExportSelection: true,
    /**
     * @cfg showExportAll {boolean} set true to show the export all button (defaults to true). 
     */
    showExportAll: true,
    /**
     * @cfg includeHidden {boolean} set to true to include hidden fields in the export (defaults to false).
     */
    includeHidden: false,
    /*
     * @cfg buttonContainerId {string/number} buttonContainerId this is an id, itemId or position of the docked item that the export buttons
     * should be added to. By default a docked item with the id/itemId of 'exportButtonContainer' is used, if that is not
     * found then the first docked item is used. Alternatively, you can specify a buttonContainer property that has a 
     * reference to the component to add the export buttons to.
     */
    buttonContainerId: 'exportButtonContainer',
    /*
     * @cfg buttonContainer {Ext.Container} see buttonContainerId config for details. 
     */
    newLine: "\n",
    init: function(cmp) {
        this.cmp = cmp;
        this.view = this.cmp.view;
        this.store = this.cmp.store;
        this.cmp.addEvents({
          'beforeexport': true,
          'export': true,
          'afterexport': true
        });
        this.cmp.on('render',this.onRender,this,{defer:200});
    },
    onRender: function(){
        if (this.buttonContainer && this.buttonContainer.add){
            this.buttonContainer.add(this.getExportButtons());
        }else{
            this.buttonContainer = this.cmp.getDockedComponent(this.buttonContainerId);
            if (!this.buttonContainer){
                this.buttonContainer = this.cmp.getDockedComponent(0);
            }
        }
        if (this.buttonContainer){
            this.buttonContainer.add(this.getExportButtons());
        }else{
            Ext.Error.raise("The Xml Export plugin could not find a docked item to add the export buttons to.");
        }
    },
    getExportButtons: function(){
        var tbb = [];
        if (this.showExportSelection){
            tbb.push({
                xtype: 'button',
                text: this.localeExportSelection,
                handler: this.exportSelection,
                scope: this
            });
        }
        if (this.showExportAll){
            tbb.push({
                xtype: 'button',
                text: this.localeExportAll,
                handler: this.exportAll,
                scope: this
            });
        }
        return tbb;
    },
    exportAll: function(){
        if (this.fireEvent('beforeexport', this, 'all') !== false){
            this.exportXmlInit(true);
        }
    },
    exportSelection: function(){
        if (this.fireEvent('beforeexport', this, 'selection') !== false){
            this.exportXmlInit(false);
        }
    },
    exportXmlInit: function(all){
        if (all){
            this.exportXml(this.store.data.items);
        }else{
            if (this.cmp.selModel.selType == 'cellmodel'){
                // cell model has a bug in 4.x
            }else{
                if (this.cmp.selModel.hasSelection()){
                    this.exportXml(this.cmp.selModel.getSelection());
                }else{
                    Ext.Msg.alert('Export','Please select items to export.');
                }
            }
        }
    },
    exportXml: function(records){
        this.pushToBrowser(this.getExcelXml(records));
    },
    pushToBrowser: function(xml){
        //document.location.href = 'data:application/vnd.ms-excel;base64,' + window.btoa(xml);
        window.open('data:application/vnd.ms-excel;base64,' + window.btoa(xml));
    },
    getExcelXml: function(records) {
        var worksheet = this.createWorksheet(records),
            totalWidth = this.cmp.headerCt.getFullWidth(),
            n = this.newLine, rec;
        return '<?xml version="1.0"?>' + n +
            '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"' + n +
            ' xmlns:o="urn:schemas-microsoft-com:office:office"' + n + 
            ' xmlns:x="urn:schemas-microsoft-com:office:excel"' + n + 
            ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"' + n +
            ' xmlns:html="http://www.w3.org/TR/REC-html40">' + n +
            ' <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">'+ n +
            '  <Title>' + this.cmp.title + '</Title>'+ n +
            '  <Version>14.0</Version>' + n +
            ' </DocumentProperties>' + n +
            ' <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">' + n +
            '   <AllowPNG/>' + n +
            ' </OfficeDocumentSettings>' + n +
            ' <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">' + n +
            '  <WindowHeight>' + worksheet.height + '</WindowHeight>' + n +
            '  <WindowWidth>' + worksheet.width + '</WindowWidth>' + n +
            '  <ProtectStructure>False</ProtectStructure>' + n +
            '  <ProtectWindows>False</ProtectWindows>' + n +
            ' </ExcelWorkbook>' + n +
            ' <Styles>' + n +
            '  <Style ss:ID="Hyperlink" ss:Name="Hyperlink">' + n +
            '   <Font ss:FontName="Helvetica" ss:Size="12" ss:Color="#0000D4" ss:Underline="Single"/>' + n +
            '  </Style>' + n +
            '  <Style ss:ID="Default">' + n +
            '   <Font ss:FontName="Helvetica" ss:Size="12" ss:Color="#000000"/>' + n +
            '  </Style>' + n +
            ' </Styles>' + n +
              worksheet.xml + n +
            '</Workbook>';
    },

    createWorksheet: function(records) {

        // Calculate cell data types and extra class names which affect formatting
        var cellType = [], 
            cellTypeClass = [],
            cm = this.cmp.headerCt,
            totalWidthInPixels = 0,
            colXml = '',
            headerXml = '',
            i, w, fld,
            visibleColumnCount,
            result, t, cellClass,
            r, k ,v, j, n = this.newLine, col;
        for (i = 0; i < cm.getColumnCount(); i++) {
            if (this.includeHidden || !cm.isHidden(i)) {
                w = cm.items.get(i).getWidth();
                totalWidthInPixels += w;
                if (cm.items.get(i).dataIndex){
                    fld = this.store.model.prototype.fields.map[cm.items.get(i).dataIndex];
                    
                    switch(fld.type.type) {
                        case "int":
                            cellType.push("Number");
                            cellTypeClass.push("int");
                            break;
                        case "float":
                            cellType.push("Number");
                            cellTypeClass.push("float");
                            break;
                        case "bool":
                        case "boolean":
                            cellType.push("String");
                            cellTypeClass.push("");
                            break;
                        case "date":
                            cellType.push("DateTime");
                            cellTypeClass.push("date");
                            break;
                        default:
                            cellType.push("String");
                            cellTypeClass.push("");
                            break;
                    }
                }
            }
        }
        visibleColumnCount = cellType.length;

        result = {
            height: 9000,
            width: Math.floor(totalWidthInPixels * 30) + 50
        };

        // Generate worksheet header details.
        t = ' <Worksheet ss:Name="' + this.cmp.title + '">' + n +
            '  <Table x:FullRows="1" x:FullColumns="1"' + n +
            '   ss:ExpandedColumnCount="' + visibleColumnCount + '"' + n + 
            '   ss:ExpandedRowCount="' + (this.store.getCount() + 2) + '">' + n;

        // Generate the data rows from the data in the Store
        for (i = 0, it = records, l = it.length; i < l; i++) {
            t += '   <Row>' + n;
            cellClass = (i & 1) ? 'odd' : 'even';
            r = it[i].data;
            k = 0;
            for (j = 0; j < cm.getColumnCount(); j++) {
                if (this.includeHidden || !cm.isHidden(j)) {
                    if (cellType[k]){
                        col = cm.items.get(j);
                        v = r[col.dataIndex];
                        rec = this.store.getAt(j);
                        if (col.xmlRenderer){
                            t += '    '+col.xmlRenderer(v,undefined,rec)+n;
                        }else{
                            t += '    <Cell ss:StyleID="Default"><Data ss:Type="' + cellType[k] + '">';
                            if (cellType[k] == 'DateTime') {
                                t += Ext.Date.format(v,'Y-m-d');
                            } else {
                                t += Ext.htmlEncode(v);
                            }
                            t +='</Data></Cell>'+ n;
                        }
                    }
                    k++;
                }
            }
            t += '    </Row>' + n;
        }

        result.xml = t + '  </Table>' +  n +
        ' </Worksheet>' + n;
        return result;
    }
});