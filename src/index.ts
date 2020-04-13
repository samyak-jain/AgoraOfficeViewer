import * as AgoraRTM from 'agora-rtm-sdk';

type RTM = typeof AgoraRTM;
type RtmClient = ReturnType<RTM['createInstance']>;
type RtmChannel = ReturnType<RtmClient['createChannel']>;

export enum Role {
    Broadcaster,
    Receiver
}

export default class AgoraOffice {
    baseUrl: string;
    channel: RtmChannel;
    iframe: HTMLIFrameElement;
    role: Role;
    fileUrl: URL;
    proxyBase: string;

    constructor(baseUrl: string, channel: RtmChannel, iframe: HTMLIFrameElement, role: Role) {
        this.baseUrl = baseUrl;
        this.channel = channel;
        this.iframe = iframe;
        this.role = role;
    }

    async renderFilePickerUi() {
        const requestUrl = new URL("/upload_ui", this.baseUrl);
        try {
            const response = await fetch(requestUrl.toString());

            if (response.ok) {
                const textResponse = await response.text();
                this.loadIframe(textResponse);

                return new Promise(resolve => {
                    window.document.addEventListener('iframeEvent', (iframeResponse: CustomEvent) => {
                        console.log("parent");
                        console.log(iframeResponse);
                        const fileUrl: URL = new URL(`/files/${iframeResponse.detail}`, this.baseUrl);
                        resolve(fileUrl);
                    })
                });
            }
        } catch (error) {
            return Promise.reject(error);
        }
    }

    async loadDocument(fileUrl: URL) {
        this.fileUrl = fileUrl;
        const requestUrl = new URL(`/do?url=${fileUrl.toString()}`, this.baseUrl);
        try {
            const response = await fetch(requestUrl.toString());

            if (response.ok) {
                const jsonResponse = await response.json();
                console.log(jsonResponse);
                this.proxyBase = jsonResponse.baseUrl;
                this.loadIframe(jsonResponse.htmlContent);
            }

            return new Promise(resolve => {
                // const parent = document.getElementsByClassName("WACFrameWord")[0];
                const doc = this.iframe.contentDocument;
                const observer = new MutationObserver(mutations => {
                    if (doc.getElementById("WACScroller") || doc.getElementById("AppForOfficeOverlay")) {
                        observer.disconnect();
                        resolve();
                    }
                });
                observer.observe(doc, {attributes: false, childList: true, characterData: false, subtree:true});
            })

        } catch (error) {
            return Promise.reject(error);
        }
    }

    syncDocument() {
        this.channel.join().then(() => {
            this._syncPPT();
        }).catch(error => {
            console.error("Error Joining RTM channel " + error,)
        });
    }

    _syncDOC() {
        if (this.role === Role.Broadcaster) {
            const scrollElement = this.iframe.contentDocument.getElementById("WACContainer");
            scrollElement.onwheel = () => {
                console.log(JSON.stringify({
                    sT: scrollElement.scrollTop,
                    sL: scrollElement.scrollLeft,
                    fileUrl: this.fileUrl
                }));
                this.channel.sendMessage({
                    text: JSON.stringify({
                        sT: scrollElement.scrollTop,
                        sL: scrollElement.scrollLeft,
                        fileUrl: this.fileUrl
                    })
                }).catch(error => {
                    console.error("There was an error trying to send a message in RTM " + error);
                })
             };
        } else if (this.role === Role.Receiver) {
            this.channel.on("ChannelMessage", ({ text }, _senderId) => {
                const scrollElement = this.iframe.contentDocument.getElementById("WACContainer");
                console.log(text);
                const scrollPosition = JSON.parse(text);
                if (this.fileUrl == undefined) {
                    this.loadDocument(scrollPosition.fileUrl);
                }
                scrollElement.scrollTo(scrollPosition.sL, scrollPosition.sT);
            });
        } else {
            console.error("Invalid Role Specified")
        }
    }

    _syncPPT() {
        if (this.role == Role.Broadcaster) {
            setInterval(() => {
                const slideNumber = this.iframe.contentDocument.querySelector("#ButtonSlideMenu-Medium14 > span").textContent.split(" ")[1];
                this.channel.sendMessage({
                    text: JSON.stringify({
                        num: slideNumber,
                        fileUrl: this.fileUrl
                    })
                });
            }, 1000);
        } else if (this.role == Role.Receiver) {
            this.channel.on("ChannelMessage", ({ text }, _senderId) => {
                const response = JSON.parse(text);
                if (this.fileUrl == undefined) {
                    this.loadDocument(response.fileUrl);
                }
                const slideNum = parseInt(response.num);
                let currentNum = parseInt(this.iframe.contentDocument.querySelector("#ButtonSlideMenu-Medium14 > span").textContent.split(" ")[1]);
                const left: HTMLDivElement = this.iframe.contentDocument.querySelector("#cell_0");
                const right: HTMLDivElement = this.iframe.contentDocument.querySelector("#cell_10");
                const leftBound = slideNum - (slideNum % 10 ? slideNum % 10 : slideNum) + 1;
                const rightBound = leftBound + 9;
                while (currentNum < leftBound || currentNum > rightBound) {
                    if (currentNum > rightBound) {
                        left.click();
                        currentNum -= 10;
                    }
                    if (currentNum < leftBound) {
                        right.click();
                        currentNum += 10;
                    }
                }
                const moveSlide = slideNum < 11 ? slideNum - 1 : slideNum % 10;
                const moveToSlide: HTMLDivElement = this.iframe.contentDocument.querySelector(`#cell_${moveSlide}`);
                moveToSlide.click();
            });
        } else {
            console.error("Invalid Role Specified");
        }
    }

    loadIframe(content: string) {
        const document = this.iframe.contentWindow.document; 
        const dom = new DOMParser().parseFromString(content, "text/html");
        const mutScript = document.createElement("script");
        const mutScriptContent = ` 
            Object.defineProperty(HTMLImageElement.prototype, 'src', {
                enumerable: true,
                get: function() {
                    return this.getAttribute('src');
                },
                set: function(url) {
                    if (!url.startsWith("http")) {
                        this.crossOrigin = '';
                        this.setAttribute('src', url);
                    }
                    else if (!url.startsWith("blob:") && !url.startsWith("data:")) {
                        // Set if not already set
                        if (this.crossOrigin !== undefined) {
                            this.crossOrigin = '';
                        }
                        const currentUrl = new URL(url);
                        if (currentUrl) {
                            const newUrl = new URL("/proxy/" + encodeURIComponent(currentUrl.host) + currentUrl.pathname, "${this.baseUrl.toString()}").toString();
                            this.setAttribute('src', newUrl);
                        } else {
                            // Set the original attribute
                            this.setAttribute('src', url);
                        }
                    } else {
                        this.crossOrigin = undefined;
                        this.setAttribute('src', url);
                    }
                },
            });
            
            // const img_observer = new MutationObserver(mutations => {
            //     mutations.forEach(mutation => {
            //         mutation.addedNodes.forEach(node => {
            //             // console.log(node);
            //             // console.log(node.tagName);
            //             if (node.tagName == "IMG") {
            //                 const currentUrl = new URL(node.src);
            //                 if (currentUrl) {
            //                     console.log(currentUrl);
            //                     const newUrl = new URL(currentUrl.pathname, "${this.baseUrl.toString()}/proxy/" + encodeURIComponent(currentUrl.host));
            //                     node.src = newUrl.toString();
            //                     node.crossOrigin = "";
            //                 }
            //             }
            //         })
            //     });  
            // });
            // img_observer.observe(document, {attributes: false, childList: true, characterData: false, subtree:true});
        `;
        mutScript.src = "data:text/javascript;charset=utf-8," + escape(mutScriptContent);
        const base = document.createElement("base");
        const url = new URL(this.baseUrl);;
        const proxyComponent = (this.proxyBase === undefined) ? "" : `proxy/${encodeURIComponent(this.proxyBase)}/`;
        base.setAttribute("href", `${url.protocol}//${url.hostname}${url.port ? ":" + url.port : ""}/${proxyComponent}`);
        const scriptTag: HTMLScriptElement = dom.querySelector("#applicationOuterContainer > script:nth-child(6)");
        console.log(scriptTag);
        if (scriptTag) {
            fetch(`${this.baseUrl.toString()}/assets/BootView.js`).then(response => {
                return response.text();
            }).then(data => {
                scriptTag.parentElement.removeChild(scriptTag);
                const newScriptTag = document.createElement("script");
                const scriptContent = document.createTextNode(data);
                newScriptTag.appendChild(scriptContent);
                dom.body.appendChild(newScriptTag);
                dom.head.insertBefore(base, dom.head.firstChild);
                dom.head.insertBefore(mutScript, dom.head.firstChild);
                document.open();
                document.write(dom.documentElement.innerHTML);
                document.close();
            });
        } else {
            dom.head.insertBefore(base, dom.head.firstChild);
            dom.head.insertBefore(mutScript, dom.head.firstChild);
            document.open();
            document.write(dom.documentElement.innerHTML);
            document.close();
        }
    }

    stopSync() {
        this.channel.leave();
    }
}