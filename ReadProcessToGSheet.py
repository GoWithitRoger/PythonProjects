import os, zipfile, re, urllib, urllib2, ntpath

path_to_directory = raw_input("Type path to directory on .docx files: ")
spreadsheet_name = raw_input("Type a name for the spreadsheet (or use existing name):")

files = [os.path.join(path_to_directory, fn) for fn in next(os.walk(path_to_directory))[2]]

for filePath in files:
    try:
        docx = zipfile.ZipFile(filePath)
    except zipfile.BadZipfile:
        continue
    docx = zipfile.ZipFile(filePath)
    content = docx.read('word/document.xml')
    cleanedtext = re.sub('<(.|\n)*?>', '', content)

    ############################# Begin Calais process
    print "\n\n\n----------- Starting Calais Processing"
    myCalaisAPI_key = 'YOUR CALAIS API KEY GOES HERE' ### enter your Calais API key before running this script
    calaisREST_URL = 'http://api.opencalais.com/enlighten/rest/' # this is the older REST interface; might not be working anymore | newer one: http://www.opencalais.com/documentation/calais-web-service-api/api-invocation/rest
    
    # alert user and shut down if the API key variable is still null.
    if myCalaisAPI_key == '':
        print "You need to set your Calais API key in the 'myCalaisAPI_key' variable, in the script before running"

    ##### set XML parameters for Calais. | see "Input Parameters" at: http://www.opencalais.com/documentation/calais-web-service-api/forming-api-calls/input-parameters
    calaisParams = '''
    <c:params xmlns:c="http://s.opencalais.com/1/pred/" xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
    <c:processingDirectives
        c:contentType="text/txt"
        c:enableMetadataType="GenericRelations"
        c:outputFormat="text/simple"/>
    <c:userDirectives
        c:allowDistribution="false"
        c:allowSearch="false" />
    <c:externalMetadata/>
    </c:params>
    '''
    
    ##### send data to Calais API. | see: http://www.opencalais.com/APICalls
    dataToSend = urllib.urlencode({
        'licenseID': myCalaisAPI_key,
        'content': cleanedtext,
        'paramsXML': calaisParams
    })

    ##### get API results
    calaisresults = urllib2.urlopen(calaisREST_URL, dataToSend).read()
    
    ############################ Trim Calais sesults
    start = calaisresults.find('--><!--') + 7
    end = calaisresults.find('--><Open', start)
    calaistrimmed = calaisresults[start:end]
    print "----------- Calais processing completed; results returned and trimmed"
    
    ############################# Begin the SheetSync process (https://sheetsync.readthedocs.io/en/latest/)
    print "Starting SheetSync injection to Google Spreadsheet"
    import sheetsync
    gUsername = "EMAIL ADDRESS GOES HERE" ### enter your gmail/Google Sheets email address before running this script
    key = ntpath.basename(filePath)
    
    target = sheetsync.Sheet(username=gUsername,
                            password="PASSWORD GOES HERE", ### enter your gmail/Google Sheets password before running this script (this may need to a an application-specific password if two-factor authentication is enabled on the Google account your's using)
                            document_name=spreadsheet_name,
                            key_column_headers=["File Name"])
    
    ### creates 'key' variable for use later, when it must be added to create a "dictionary within a dictionary"
    key = '{"'+key+'": '
    ### removing spaces, adding quotes, and line breaks
    sNospace = '":'.join(calaisresults[start:end].split(": ",-1))
    sCollon = ': "'.join(sNospace.split(":",-1))
    sComma = ', '.join(sCollon.split(",",-1))
    sRquotespace = '"'.join(sComma.split(' "',-1))
    sLastquote = """"},}""".join(sRquotespace.rsplit(',',1))
    sLbracket = '\n{"'.join(sLastquote.split("\n",1))
    sEndquote = '" \n, '.join(sLbracket.split(", \n",-1))
    sBeginquote = ', "'.join(sEndquote.split("\n, ",-1))
    sFormatted = sBeginquote
    sFixorigin = ', Origin '.join(sFormatted.split("\nOrigin\n",-1))
    sFixHTTP = 'http:'.join(sFixorigin.split('http:"',-1))
    
    sPluskey = key + sFixHTTP
    sFinal = ': '.join(sPluskey.split(": \n",1))
    print "\n\n#### Results formatted for SheetSync #### \n\n" + sFinal
    
    ### uses eval to create the string from above to a dictionary
    dict_final = eval(sFinal)
    data = dict_final
    
    ### 'injects' dictionary
    target.inject(data)
    print "\n----------- Success! SheetSycn process completed\n##### Review the new spreadsheet created here: ####\n-----------------------------------------------------------\n %s" % target.document_href
