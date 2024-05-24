Class WS
{
    __new(url,Callback:= 0)
    {
        this.Client := WebSocket(url,{message: (self, data) => this.Reponse(data)})
        this.Callback := Callback ; callbackfunction
    }

    Send(Method,Params:=Map())
    {
        static id := 1
        this.Client.sendText(
            JSON.stringify(
                Map(
                    "id",++id, 
                    "method",Method,
                    "params",Params
                )
            )
        )
    }
    /*
     Note: reponse receives unparsed json data
     reason AHK ins single threaded and we can't use async/await to parse the data
     incase data from callback will being parsed than AHK is busy doing parsing and can't receive new data
     from websocket so callback function will misses the data
     this is why user has to parse the data with out disturbing callback function
    */
    Reponse(data) 
    {
        if this.Callback
            this.Callback.Call(data) ; passing data to callback function
    }

    static observe := [
        'Page.domContentEventFired',
        'Page.fileChooserOpened',
        'Page.frameAttached',
        'Page.frameDetached',
        'Page.frameNavigated',
        'Page.interstitialHidden',
        'Page.interstitialShown',
        'Page.javascriptDialogClosed',
        'Page.javascriptDialogOpening',
        'Page.lifecycleEvent',
        'Page.loadEventFired',
        'Page.windowOpen',
        'Page.frameClearedScheduledNavigation',
        'Page.frameScheduledNavigation',
        'Page.compilationCacheProduced',
        'Page.downloadProgress',
        'Page.downloadWillBegin',
        'Page.frameRequestedNavigation',
        'Page.frameResized',
        'Page.frameStartedLoading',
        'Page.frameStoppedLoading',
        'Page.navigatedWithinDocument',
        'Page.screencastFrame',
        'Page.screencastVisibilityChanged',
        'Network.dataReceived',
        'Network.eventSourceMessageReceived',
        'Network.loadingFailed',
        'Network.loadingFinished',
        'Network.requestServedFromCache',
        'Network.requestWillBeSent',
        'Network.responseReceived',
        'Network.webSocketClosed',
        'Network.webSocketCreated',
        'Network.webSocketFrameError',
        'Network.webSocketFrameReceived',
        'Network.webSocketFrameSent',
        'Network.webSocketHandshakeResponseReceived',
        'Network.webSocketWillSendHandshakeRequest',
        'Network.requestWillBeSentExtraInfo',
        'Network.resourceChangedPriority',
        'Network.responseReceivedExtraInfo',
        'Network.signedExchangeReceived',
        'Network.requestIntercepted'
      ]
}
