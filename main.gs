var doc_id = '1rl45vS425DC9cE85A4xnBoJUjqyDl_0zd8ksj9QU1BA';
var doc = DocumentApp.openById(doc_id);
var ss_id = '0AsBD2EanwzGQdFRXTG5nbDA5RklnMDNvTzJ3U3JpcWc';
var ss = SpreadsheetApp.getActiveSpreadsheet();

function main() {
  var pubmed_urls = getPbUrls();
  for (i=0;i<pubmed_urls.length;i++){
    var pmid = getPMID(pubmed_urls[i].url);
    var xml = getCitationXml(pmid);
    var cit_obj = getCitationObj(pubmed_urls[i], xml);
    writeInLineCitations(cit_obj);
    if (!referenceExists(cit_obj)){
      putInSS(cit_obj);
    }
  }
  writeBibliography();
}

function getPbUrls() {
  var text = doc.getText();
  var index = 0;
  var i = 0;
  var pubmed_urls = new Array();
  while (true){
    index = text.indexOf('http\://www\.ncbi\.nlm\.nih\.gov/pubmed/',index);
    if (index == -1){break;}
    else{
      pubmed_urls[i] = {"url":text.slice(index,index+43), "index":index};
      i++;
    }
    index++;
  }
return pubmed_urls;
}

function getPMID(pubmed_url){
  var patt = /pubmed\/\d+/;
  var pmid = pubmed_url.match(patt);
  if (pmid==-1){Browser.msgBox('pmid not found');}
  return pmid;
}

function getCitationXml(pmid) {
  var base_url = 'http://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?';
  var query = 'db=pubmed&retmode=xml&id='+pmid;
  var response =UrlFetchApp.fetch(base_url+query);
  var text = response.getContentText();
  var xml_doc = Xml.parse(text, true);
  var set = xml_doc.getElement();
  var article = set.getElement();      
  var citation = article.getElement();
  var title = citation.getElement('Article').getElement('ArticleTitle').getText();
  return citation;
}

function getCitationObj(pburl, medlinecitation){
  var pmid = medlinecitation.getElement('PMID').getText();
  var article = medlinecitation.getElement("Article");
  var authorlist = article.getElement("AuthorList");
  var authorlistarray = authorlist.getElements();
  var numauthors = authorlistarray.length;
  end = numauthors;
  var etAl = false;
  if (numauthors > 2)  {    etAl = true;  }
  var authorlistValue = "";
  var inlineAuthorlistValue = "";
  var i = 0;
  for (i=0;i<=end-1;i++)
  {
      var author = authorlistarray[i];
      var lastname = author.getElement("LastName");
      var initials = author.getElement("Initials");
      var lastnameValue = lastname.getText();
      var initialsValue = initials.getText();
      if (i==0){
          inlineAuthorlistValue += lastnameValue;
      }
      if (i==1) 
      {
         if (etAl) {
             inlineAuthorlistValue += ' et al.';
         }
         else {
             inlineAuthorlistValue += ' and '+lastnameValue;
         }
              
      }
      authorlistValue += lastnameValue + " " + initialsValue;
      if(i<end-1) {authorlistValue += ", ";}
  }
    var articleTitle = article.getElement("ArticleTitle").getText();
  var pagination = article.getElement("Pagination");
  var medlinepgn = pagination.getElement("MedlinePgn").getText();                               
  var journal = article.getElement("Journal");
  var journalissue = journal.getElement("JournalIssue");
  var journalName = journal.getElement("ISOAbbreviation").getText();
  var volume = journalissue.getElement("Volume");
  var volumeValue = volume.getText();
  var pubdate = journalissue.getElement("PubDate");
  var year = pubdate.getElement("Year").getText();
  return { 'authors':authorlistValue, 'year':year, 'title':articleTitle, 'journal':journalName, 'volume':volumeValue, 'pages':medlinepgn, 'pmid':pmid, 'inLineAuthors':inlineAuthorlistValue, 'url':pburl.url };
}

function referenceExists(citation){
  var pmids = ss.getActiveSheet().getRange(1, 7,ss.getLastRow()+1).getValues();
  for (var x=0;x<pmids.length;x++){
    if (pmids[x]==citation.pmid){
      return true;
    }
  }   
  return false;
}

function putInSS (citation){
  var lastrow = ss.getLastRow();
  var cell = ss.getRange('a'+(Number(ss.getLastRow())+1))
  var row = 0;
  var col = 0;
  for (var key in citation) {
    cell.offset(row, col).setValue(citation[key]).setBackground('aqua');
    col++;
  }
  ss.sort(1);
}

function writeInLineCitations(citation_obj){
  var inline_citation = '('+citation_obj.inLineAuthors+', '+citation_obj.year+')';
  doc.editAsText().replaceText('http\://www\.ncbi\.nlm\.nih\.gov/pubmed/'+citation_obj.pmid, inline_citation);
}

function writeBibliography(){
  removeBibliography();
  doc.appendParagraph('References').setBold(true);
  var references = ss.getActiveSheet().getRange(1,1,ss.getLastRow(),ss.getLastColumn()).getValues();
  for (var i=0;i<references.length;i++) {
    var bib_entry = doc.appendParagraph(references[i][0]).setFontSize(10).setLineSpacing(1).setSpacingAfter(1).setSpacingBefore(1).setBold(false);
    bib_entry.appendText(' ('+references[i][1]+')');
    bib_entry.appendText(' '+references[i][2]).setItalic(true);
    bib_entry.appendText(' '+references[i][3]).setBold(true).setItalic(false);
    bib_entry.appendText(' '+references[i][4]).setBold(false);
    bib_entry.appendText(': '+references[i][5]);
    bib_entry.appendText(' PMID '+references[i][6]).setLinkUrl(references[i][8]).setFontSize(8);
  }
  
}

function removeBibliography(){
  var body = doc.getBody();
  for ( var i = 0; i < body.getNumChildren(); i ++)
  {
    var child = body.getChild(i);
    if (child.getText() == 'References')
      Logger.log(i + child.getNextSibling().getText());
    
  }
 }
  
}

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [ {name: "Update citations", functionName: "main"} ];
  ss.addMenu("Pubmed cite", menuEntries);
}

function onInstall() {
  onOpen();
}
