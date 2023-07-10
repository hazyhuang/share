/**
 * Purchase Order Report
 *
 * @NScriptName Purchase Order PDF Report
 * @NScriptType UserEventScript
 * @NApiVersion 2.1
 */
define(["N/log", "N/email", "N/url", "N/search", "N/record", "N/file", "N/render", "../util/commonsHelper", "./contact_dao"], function (log, email, url, search, record, file, render, commons, contact_dao) {
    var exports = {};
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.createPDF = exports.loadData = exports.afterSubmit = exports.beforeLoad = void 0;
    function getVendorContact(vendor_id) {
        var contacts = contact_dao.fetchData({ vendor_internal_id: vendor_id });
        if (contacts.length > 0) {
            return contacts[0].name;
        }
        else {
            return "";
        }
    }
    function beforeLoad(context) {
        if (context.type == context.UserEventType.VIEW) {
            var rec_record = context.newRecord;
            var rec_id = rec_record.getValue({
                fieldId: 'id'
            });
            var output = url.resolveScript({
                scriptId: 'customscript_purchase_order_suitelet',
                deploymentId: 'customdeploy_purchase_order_suitelet',
                returnExternalUrl: false
            });
            output = output + '&internalid=' + rec_id;
            log.debug("output", output);
            var form = context.form;
            form.addButton({
                id: 'custpage_isbutton',
                label: '采购單列印',
                functionName: 'window.open( "' + output + '","_parent")'
            });
        }
    }
    exports.beforeLoad = beforeLoad;
    //修改PO状态，触发发送报表动作
    function afterSubmit(context) {
        if (context.type == context.UserEventType.EDIT) {
            var newRecord = context.newRecord;
            var oldRecord = context.oldRecord;
            var newStatus = newRecord.getValue({ fieldId: 'approvalstatus' });
            var oldStatus = oldRecord.getValue({ fieldId: 'approvalstatus' });
            if (newStatus != oldStatus && newStatus == "2") { // 状态2为 Approved
                log.debug('oldstatus', oldStatus);
                sendReport(newRecord);
            }
        }
    }
    exports.afterSubmit = afterSubmit;
    function getShipToBillTo(purchase_order, sub_record, template_type) {
        var shipto = "";
        var sub_phone = "";
        if (template_type == "tw") {
            shipto = purchase_order.getValue({ fieldId: 'custbody_xxadj0073' });
            if (!commons.makesure(shipto)) {
                shipto = purchase_order.getValue({ fieldId: 'shipaddress' });
            }
        }
        else if (template_type == "cn") {
            var shipto_key = purchase_order.getValue({ fieldId: 'shippingaddress_key' });
            log.audit('shipto_key', shipto_key);
            if (commons.makesure(shipto_key)) {
                var loc_record = record.load({ type: 'address', id: shipto_key });
                var addressee = loc_record.getValue({ fieldId: 'addressee' });
                var addr1 = loc_record.getValue({ fieldId: 'addr1' });
                if (commons.makesure(addressee)) {
                    shipto = addressee + '<br/>' + addr1;
                    log.audit('shipto', shipto);
                }
            }
            if (!commons.makesure(shipto)) {
                shipto = sub_record.getValue({ fieldId: 'shippingaddress_text' });
            }
            sub_phone = sub_record.getValue({ fieldId: 'fax' });
        }
        var billto_id = purchase_order.getValue({ fieldId: 'custbody_xxadj0074' });
        var billto = "";
        if (commons.makesure(billto_id)) {
            var addr_record = record.load({ type: 'customrecord_xxadj0113', id: billto_id });
            if (commons.makesure(addr_record)) {
                billto = addr_record.getValue({ fieldId: "custrecord_xxadj0113_billing_address" });
                billto = billto.replace(new RegExp("\n", "gm"), "<br/>");
            }
        }
        if (billto == "") {
            billto = purchase_order.getValue({ fieldId: 'billaddress' });
        }
        return {
            billto: billto,
            shipto: shipto,
            sub_phone: sub_phone
        };
    }
    //加载数据
    function loadData(rec) {
        var vendor_id = rec.getValue({ fieldId: 'entity' });
        var employee_id = rec.getValue({ fieldId: 'employee' });
        var approver_id = rec.getValue({ fieldId: 'custbody_xxadj0092' });
        var approver_record;
        if (String(approver_id) != "") {
            approver_record = record.load({
                type: record.Type.EMPLOYEE,
                id: approver_id
            });
        }
        var employee_record;
        var user_name = "";
        if (String(employee_id) != "") {
            employee_record = record.load({
                type: record.Type.EMPLOYEE,
                id: employee_id
            });
            user_name = employee_record.getValue({ fieldId: 'firstname' });
        }
        var not_send_auto = rec.getValue({ fieldId: 'custbody_xxadj0088' });
        //Send Notes
        var sendNotes = rec.getValue({ fieldId: 'custbody_xxadj0107' });
        var vendor_record = record.load({ type: record.Type.VENDOR, id: vendor_id });
        var vendor_name = vendor_record.getValue({ fieldId: 'altname' });
        var contacts = getContact(getContactSearch(rec.id));
        var sub_id = rec.getValue({ fieldId: 'subsidiary' });
        var subsidiary_record = record.load({ type: record.Type.SUBSIDIARY, id: sub_id });
        var subsidiary_name = subsidiary_record.getValue({ fieldId: 'name' });
        var email_body = subsidiary_record.getValue({ fieldId: 'custrecord_xxadj0102' });
        log.audit('email_body', email_body);
        var cc_emails = subsidiary_record.getValue({ fieldId: 'custrecord_xxadj0106' });
        var email_attach = null;
        if (sendNotes) {
            var attachEmail = subsidiary_record.getValue({ fieldId: 'custrecord_xxadj0103' });
            email_attach = file.load({ id: attachEmail });
        }
        var doc_number = rec.getValue({ fieldId: 'tranid' });
        log.debug('approver_id', approver_id);
        log.debug('approver_record', approver_record);
        var template = getTemplateFile(subsidiary_name);
        var addr = getShipToBillTo(rec, subsidiary_record, template.template_type);
        var project_code = getProjectCode(rec);
        return {
            not_send_auto: not_send_auto,
            vendor_record: vendor_record,
            vendor_name: vendor_name,
            subsidiary_record: subsidiary_record,
            subsidiary_name: subsidiary_name,
            contacts: contacts,
            email_body: email_body,
            cc_emails: cc_emails,
            email_attach: email_attach,
            doc_number: doc_number,
            user_name: user_name,
            employee_record: employee_record,
            approver_record: approver_record,
            approver_id: approver_id,
            employee_id: employee_id,
            xmlTemplateFile: template.xmlTemplateFile,
            template_type: template.template_type,
            shipto: addr.shipto,
            billto: addr.billto,
            sub_phone: addr.sub_phone,
            project_code: project_code,
            tax_rate: getTaxRate(rec),
            buyer: getEmployee(rec),
            vendor_contact: getVendorContact(vendor_record.id)
        };
    }
    exports.loadData = loadData;
    function getProjectCode(purchase_order) {
        var project_code = "";
        var lineCount = purchase_order.getLineCount({ sublistId: 'item' });
        if (lineCount >= 1) {
            var project_desc = purchase_order.getSublistText({ sublistId: 'item', fieldId: 'customer_display', line: 0 });
            // let project_desc_value = purchase_order.getSublistValue({ sublistId: 'item', fieldId: 'customer_display', line: 0 }) as string;
            if (commons.makesure(project_desc)) {
                if (project_desc.length >= 6) {
                    var project_array = project_desc.split(' ');
                    if (commons.makesure(project_array[0])) {
                        if (project_array[0].length > 2) {
                            //log.audit('project_desc_value',project_array[0]);
                            project_code = project_array[0];
                        }
                    }
                }
            }
        }
        return project_code;
    }
    function getTaxRate(purchase_order) {
        var tax_rate_ret = "";
        var lineCount = purchase_order.getLineCount({ sublistId: 'item' });
        if (lineCount >= 1) {
            var tax_rate = purchase_order.getSublistText({ sublistId: 'item', fieldId: 'taxrate1', line: 0 });
            if (commons.makesure(tax_rate)) {
                tax_rate_ret = tax_rate;
            }
        }
        return tax_rate_ret;
    }
    function getEmployee(purchase_order) {
        var emp_text = purchase_order.getText({ fieldId: 'employee' });
        return clearName(emp_text);
    }
    function clearName(name) {
        if (commons.makesure(name)) {
            var indexstart = name.indexOf('-');
            if (indexstart != -1) {
                var name_array = name.split('-');
                if (commons.makesure(name_array[1])) {
                    return name_array[1];
                }
            }
            return name;
        }
        return '';
    }
    function getTemplateFile(subsidiary_name) {
        var SUBSIDIARY = ['友達智匯智能製造(蘇州)', '艾聚達信息技術(蘇州)', '宇沛環保科技(山東)', '友達智匯智能製造(廈門)', '友達數位科技服務(蘇州)', 'AUO Digitech Pte. Ltd', '黑龙江天乐达智能显示科技有限公司', '艾杰达人工智能科技（苏州）']; //'友達數位科技服務',
        //let SUBSIDIARY = ['友達智匯智能製造(蘇州)', '艾聚達信息技術(蘇州)', '友達宇沛永續科技(蘇州)', '宇沛環保科技(山東)', '友達頤康信息科技(蘇州)', '友達智匯智能製造(廈門)', '友達數位科技服務(蘇州)'];
        var xmlTemplateFile;
        log.debug('subsidary_name', subsidiary_name);
        var template_type;
        if (SUBSIDIARY.indexOf(subsidiary_name) != -1) {
            log.debug('LANG', 'CN_CHINESE');
            xmlTemplateFile = file.load('Templates/PDF Templates/purchase_order.template_cn.xml');
            template_type = "cn";
        }
        else {
            log.debug('LANG', 'TW_CHINESE');
            xmlTemplateFile = file.load('Templates/PDF Templates/purchase_order.template_tw.xml');
            template_type = "tw";
        }
        return {
            xmlTemplateFile: xmlTemplateFile,
            template_type: template_type
        };
    }
    function renderAddData(renderer, rec, dataVO, reportDataStr) {
        renderer.addRecord({
            templateName: 'record',
            record: rec
        });
        renderer.addRecord({
            templateName: 'vendor',
            record: dataVO.vendor_record
        });
        renderer.addRecord({
            templateName: 'subsidiary',
            record: dataVO.subsidiary_record
        });
        if (String(dataVO.employee_id) != "") {
            renderer.addRecord({
                templateName: 'employee',
                record: dataVO.employee_record
            });
        }
        if (String(dataVO.approver_id) != "") {
            renderer.addRecord({
                templateName: 'approver',
                record: dataVO.approver_record
            });
        }
        renderer.addCustomDataSource({
            format: render.DataSource.JSON,
            alias: "JSON_STR",
            data: reportDataStr
        });
        var jsonObj = JSON.parse(reportDataStr);
        log.debug('jsonObj', jsonObj);
        renderer.addCustomDataSource({
            format: render.DataSource.OBJECT,
            alias: "jsonObj",
            data: jsonObj
        });
        return renderer;
    }
    function createPDF(rec, dataVO) {
        //let template = getTemplateFile(dataVO.subsidiary_name);
        var start_time = (new Date()).getTime();
        var renderer = render.create();
        renderer.templateContent = dataVO.xmlTemplateFile.getContents();
        var reportData = exportXML(dataVO, rec, dataVO.contacts);
        var reportDataStr = JSON.stringify(reportData);
        var str_length = reportDataStr.length;
        reportDataStr = reportDataStr.substring(1, str_length - 1);
        var end_time = (new Date()).getTime();
        log.audit("init renderer ms:", end_time - start_time);
        start_time = (new Date()).getTime();
        renderer = renderAddData(renderer, rec, dataVO, reportDataStr);
        end_time = (new Date()).getTime();
        log.audit("renderAddData ms:", end_time - start_time);
        var renderXmlAsString = renderer.renderAsString();
        log.debug('renderXmlAsString', renderXmlAsString);
        var invoicePdf = render.xmlToPdf({
            xmlString: renderXmlAsString
        });
        invoicePdf.name = dataVO.doc_number + '.pdf';
        return invoicePdf;
    }
    exports.createPDF = createPDF;
    //发送报表启动入口
    function sendReport(rec) {
        var dataVO = loadData(rec);
        var invoicePdf = createPDF(rec, dataVO);
        var receive_emails = [];
        for (var i = 0; i < dataVO.contacts.length; i++) {
            if (dataVO.contacts[i].email != "") {
                receive_emails.push(dataVO.contacts[i].email);
            }
        }
        log.audit('receive_emails1', receive_emails);
        var ccEmailArray = dataVO.cc_emails.split(';');
        var cc_mails = [];
        if (ccEmailArray.length > 0) {
            for (var i = 0; i < ccEmailArray.length; i++) {
                cc_mails.push(ccEmailArray[i]);
            }
        }
        var emailContent = {
            company: dataVO.subsidiary_name,
            poNumber: dataVO.doc_number,
            invoicePdf: invoicePdf,
            receiveemail: receive_emails,
            cc: cc_mails,
            attach: dataVO.email_attach,
            employee_id: dataVO.employee_id,
            emailBody: dataVO.email_body,
            vendor_name: dataVO.vendor_name
        };
        if (receive_emails.length > 0 && String(dataVO.employee_id) != "" && !dataVO.not_send_auto) {
            log.debug('emailContent', emailContent);
            sendEmail(emailContent);
        }
    }
    //发送邮件
    function sendEmail(emailContent) {
        var attachments = [];
        if (commons.makesure(emailContent.attach)) {
            attachments = [emailContent.invoicePdf, emailContent.attach];
        }
        else {
            attachments = [emailContent.invoicePdf];
        }
        if (commons.makesure(emailContent.cc)) {
            email.send({
                author: emailContent.employee_id,
                recipients: emailContent.receiveemail,
                cc: emailContent.cc,
                subject: emailContent.company + " 採購訂單-" + emailContent.poNumber + " " + emailContent.vendor_name,
                body: emailContent.emailBody,
                attachments: attachments,
            });
        }
        else {
            email.send({
                author: emailContent.employee_id,
                recipients: emailContent.receiveemail,
                subject: emailContent.company + " 採購訂單-" + emailContent.poNumber + " " + emailContent.vendor_name,
                body: emailContent.emailBody,
                attachments: attachments,
            });
        }
    }
    //格式化天
    function formatDate(dateNum) {
        var numStr = "";
        if (dateNum < 10) {
            numStr = '0' + dateNum.toString();
        }
        else {
            numStr = dateNum.toString();
        }
        return numStr;
    }
    //获取今日日期
    function getToday() {
        var current_datetime = new Date();
        var monthStr = formatDate(current_datetime.getMonth() + 1);
        var formatted_date = current_datetime.getFullYear() + "/" + monthStr + "/" + formatDate(current_datetime.getDate());
        return formatted_date;
    }
    //将数据保存成模板可识别的JSON数据
    function exportXML(dataVO, purchase_order, contacts) {
        var json_list = [];
        var head_list = [];
        var item_list = [];
        var cname = "";
        var phone = "";
        var fax = "";
        if (contacts.length > 0) {
            cname = contacts[0].cname;
            phone = contacts[0].phone;
            fax = contacts[0].fax;
        }
        var createByWO = getLine(purchase_order);
        log.debug("createByWO", createByWO);
        head_list.push({
            'today': getToday(),
            'currentUser': dataVO.user_name,
            'project_code': dataVO.project_code,
            'tax_rate': dataVO.tax_rate,
            'buyer': dataVO.buyer,
            'sub_phone': dataVO.sub_phone,
            'cname': cname,
            'phone': phone,
            'fax': fax,
            'shipto': dataVO.shipto,
            'billto': dataVO.billto,
            'createdWO': createByWO,
            'AET': '友達宇沛永續科技',
            'vendor_contact': dataVO.vendor_contact
        });
        if (createByWO) { //如果是委外订单
            var lines = getLineValues(purchase_order);
            item_list = lines;
        }
        json_list.push({
            'header': head_list,
            'items': item_list,
        });
        return json_list;
    }
    //判断Purchase Order是否为委外工单
    function getLine(purchase_order) {
        var lineCount = purchase_order.getLineCount({ sublistId: 'item' });
        for (var i = 0; i < lineCount; i++) {
            var createdWO = purchase_order.getSublistValue({ sublistId: 'item', fieldId: 'createoutsourcedwo', line: i });
            log.debug("createWO", createdWO);
            if (createdWO != undefined && createdWO.length > 3) {
                return true;
            }
        }
        return false;
    }
    //获取Purchase Order的行数据
    function getLineValues(purchase_order) {
        var lines = [];
        var lineCount = purchase_order.getLineCount({ sublistId: 'item' });
        for (var i = 0; i < lineCount; i++) {
            var aitem = purchase_order.getSublistValue({ sublistId: 'item', fieldId: 'item', line: i });
            var subRecord = record.load({ type: record.Type.OTHER_CHARGE_ITEM, id: aitem });
            var description = subRecord.getValue({
                fieldId: 'purchasedescription'
            });
            var internalid = subRecord.getValue({
                fieldId: 'internalid'
            });
            var itemid = subRecord.getValue({
                fieldId: 'itemid'
            });
            lines.push({ itemid: itemid, internalid: internalid, description: description });
            log.debug("itemid", itemid);
        }
        return lines;
    }
    //获取联络人搜索
    function getContactSearch(cust_id) {
        var savedSearch = search.load({
            id: 'customsearch_po_contact_search'
        });
        savedSearch.filters = [search.createFilter({
                name: 'internalid',
                join: 'transaction',
                operator: search.Operator.IS,
                values: cust_id
            })];
        log.debug('savedSearch.column', savedSearch.columns);
        log.audit('savedSearch.filters', savedSearch.filters);
        var result = savedSearch.run().getRange({
            start: 0,
            end: 50
        });
        log.debug({
            title: 'Result',
            details: result
        });
        return result;
    }
    /*
    function getContactSearch(cust_id: any) {
        let result = search.create({
            type: 'contact',
            columns: ['entityid', 'phone', 'fax', 'email'],
            filters: [
                search.createFilter({
                    name: 'internalid',
                    join: 'vendor',
                    operator: search.Operator.IS,
                    values: cust_id
                })
            ]
        }).run().getRange({
            start: 0,
            end: 1
        });
        log.debug({
            title: 'Result',
            details: result
        });
        return result;
    
    }
    */
    /*
    function getXMLString(rawStr: string): string {
    let re = /\n/gi;
    let xmlStr = rawStr.replace(re, '<br/>');
    return xmlStr;
    
    }
    */
    //通过搜索获取 联系人
    function getContact(result) {
        var contacts = [];
        for (var i = 0; i < result.length; i++) {
            var contact = { cname: "", phone: "", fax: "", email: "", role: "" };
            var cname = result[i].getValue({ name: 'entityid' });
            var phone = result[i].getValue({ name: 'phone' });
            var mobilephone = result[i].getValue({ name: 'fax' });
            var email_1 = result[i].getValue({ name: 'email' });
            var roles = result[i].getValue({ name: 'contactrole' });
            log.audit(' cname', cname);
            log.audit(' email', email_1);
            contact.cname = cname.toString();
            contact.phone = phone.toString();
            contact.fax = mobilephone.toString();
            contact.email = email_1.toString();
            contact.role = arrayToString(roles);
            contacts.push(contact);
        }
        return contacts;
    }
    //数组转换成字符串
    function arrayToString(roles) {
        var roles_str = "";
        for (var i; i < roles.length; i++) {
            roles_str = roles_str + roles[i] + ";";
        }
        return roles_str;
    }
    return exports;
});
