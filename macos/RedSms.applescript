use framework "Foundation"

property apiBase : "https://cp.redsms.ru/api/"

-- Supply credentials and headers on stdin, so they are absent from process arguments.
on runCommand(executablePath, args, inputText)
    set commandTask to current application's NSTask's alloc()'s init()
    set inputPipe to current application's NSPipe's pipe()
    set outputPipe to current application's NSPipe's pipe()
    commandTask's setLaunchPath:executablePath
    commandTask's setArguments:args
    commandTask's setStandardInput:inputPipe
    commandTask's setStandardOutput:outputPipe
    commandTask's setStandardError:outputPipe
    commandTask's |launch|()

    set inputHandle to inputPipe's fileHandleForWriting()
    set inputString to current application's NSString's stringWithString:inputText
    inputHandle's writeData:(inputString's dataUsingEncoding:(current application's NSUTF8StringEncoding))
    inputHandle's closeFile()

    set outputData to outputPipe's fileHandleForReading()'s readDataToEndOfFile()
    commandTask's waitUntilExit()
    set outputText to (current application's NSString's alloc()'s initWithData:outputData encoding:(current application's NSUTF8StringEncoding)) as text
    if (commandTask's terminationStatus() as integer) is not 0 then error outputText
    return outputText
end runCommand

on request(parameters)
    try
        set fields to (current application's NSString's stringWithString:parameters)'s componentsSeparatedByString:linefeed
        set fields to fields as list
        set operation to item 1 of fields
        set login to item 2 of fields
        set apiKey to item 3 of fields
        if login is "" or apiKey is "" or login contains return then error "Missing or invalid RED SMS credentials"

        set curlArgs to {"--disable", "--silent", "--show-error", "--max-time", "15", "--proto", "=https", "--header", "@-", "--write-out", linefeed & "HTTP %{http_code}" & linefeed}
        set signatureValues to ""
        if operation is "info" then
            set endpoint to "client/info"
        else if operation is "send" then
            set senderName to item 4 of fields
            set recipient to item 5 of fields
            set messageText to "Тестовое сообщение от redsms.ru"
            set endpoint to "message"
            -- The signature uses the values of from, route, text, to in that order.
            set signatureValues to senderName & "sms" & messageText & recipient
            set curlArgs to curlArgs & {"--data-urlencode", "from=" & senderName, "--data-urlencode", "route=sms", "--data-urlencode", "text=" & messageText, "--data-urlencode", "to=" & recipient}
        else if operation is "status" then
            set messageId to item 4 of fields
            set uuidPattern to "[0-9A-Fa-f]{8}(-[0-9A-Fa-f]{4}){3}-[0-9A-Fa-f]{12}"
            set uuidPredicate to current application's NSPredicate's predicateWithFormat:"SELF MATCHES %@" argumentArray:{uuidPattern}
            if not (uuidPredicate's evaluateWithObject:messageId) then error "Invalid message UUID"
            set signatureValues to "uuid,status,errCode"
            set endpoint to "message/" & messageId & "?fields=uuid%2Cstatus%2CerrCode"
        else
            error "Unknown RED SMS operation"
        end if

        set ts to (current application's NSUUID's UUID()'s UUIDString()) as text
        set digest to my runCommand("/sbin/md5", {"-q"}, signatureValues & ts & apiKey)
        set signature to paragraph 1 of digest
        set headers to "login: " & login & linefeed & "ts: " & ts & linefeed & "sig: " & signature & linefeed
        return my runCommand("/usr/bin/curl", curlArgs & {"--url", apiBase & endpoint}, headers)
    on error messageText
        return "RED SMS: " & messageText
    end try
end request
