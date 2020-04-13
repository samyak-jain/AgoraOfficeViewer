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
                this.proxyBase = jsonResponse.baseUrl;
                this.loadIframe(jsonResponse.htmlContent);
            }

            return new Promise(resolve => {
                const doc = this.iframe.contentDocument;
                const observer = new MutationObserver(mutations => {
                    if (doc.getElementById("WACScroller") || doc.getElementById("AppForOfficeOverlay") || doc.querySelector("#m_excelWebRenderer_ewaCtl_m_sheetTabBar > div.ewa-stb-contentarea > div > ul")) {
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

    async _waitForDocument() {
        return new Promise(resolve => {
            this.channel.on("ChannelMessage", ({ text }, _senderId) => {
                const response = JSON.parse(text);
                resolve(response.fileUrl);
            });
        })
    }

    _sync() {
        if (this.iframe.contentDocument.getElementById("WACScroller")) {
            this._syncDOC();
        } else if (this.iframe.contentDocument.querySelector("#m_excelWebRenderer_ewaCtl_m_sheetTabBar > div.ewa-stb-contentarea > div > ul")) {
            this._syncXLS();
        } else {
            this._syncPPT();
        }
    }

    syncDocument() {
        this.channel.join().then(() => {
            if (this.role === Role.Receiver) {
                this._waitForDocument().then(fileUrl => {
                    this.loadDocument(<URL> fileUrl).then(() => {
                        this._sync();
                    });
                });
            } else {
                this._sync();
            }
        }).catch(error => {
            console.error("Error Joining RTM channel " + error,)
        });
    }

    _syncXLS() {
        const iframeDocument = this.iframe.contentDocument;
        if (this.role === Role.Broadcaster) {
            const scrollElement = iframeDocument.getElementById("m_excelWebRenderer_ewaCtl_sheetContentDiv");
            scrollElement.onwheel = () => {
                this.channel.sendMessage({
                    text: JSON.stringify({
                        type: "scroll",
                        sT: scrollElement.scrollTop,
                        sL: scrollElement.scrollLeft,
                        fileUrl: this.fileUrl
                    })
                });
            }

            const bottomList = iframeDocument.querySelector("#m_excelWebRenderer_ewaCtl_m_sheetTabBar > div.ewa-stb-contentarea > div > ul");
            for (let ele of bottomList.children) {
                ele.addEventListener('click', event => {
                    const target: HTMLSpanElement = <HTMLSpanElement> event.target;
                    const textContent = target.textContent;
                    const listBottom = iframeDocument.querySelectorAll("ul > li > a > span > span:nth-child(1)");
                    let indexClicked = 0;
                    for (let [index, ele] of listBottom.entries()) {
                        if (ele.textContent == textContent) {
                            indexClicked = index;
                        }
                    }
                    this.channel.sendMessage({
                        text: JSON.stringify({
                            type: "click",
                            index: indexClicked,
                            fileUrl: this.fileUrl
                        })
                    })
                })
            }

        } else if (this.role === Role.Receiver) {
            this.channel.on("ChannelMessage", ({ text }, _senderId) => {
                const response = JSON.parse(text);
                if (this.fileUrl == undefined) {
                    this.loadDocument(response.fileUrl);
                }
                if (response.type == "scroll") {
                    const scrollElement = iframeDocument.getElementById("m_excelWebRenderer_ewaCtl_sheetContentDiv");
                    scrollElement.scrollTo(response.sL, response.sT);
                } else if (response.type == "click") {
                    const indexClicked = response.index;
                    const listBottom = iframeDocument.querySelectorAll("ul > li > a > span > span:nth-child(1)");
                    const nodeToNavigate = listBottom[indexClicked];
                    var clickEvent = new MouseEvent('mousedown', {
                        view: window,
                        bubbles: true,
                        cancelable: true
                    });
                    nodeToNavigate.dispatchEvent(clickEvent);
                }
            });
        } else {
            console.error("Invalid Role Specified");
        }
    }

    _syncDOC() {
        if (this.role === Role.Broadcaster) {
            const scrollElement = this.iframe.contentDocument.getElementById("WACContainer");
            scrollElement.onwheel = () => {
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
        if (this.role === Role.Broadcaster) {
            setInterval(() => {
                const slideNumber = this.iframe.contentDocument.querySelector("#ButtonSlideMenu-Medium14 > span").textContent.split(" ")[1];
                this.channel.sendMessage({
                    text: JSON.stringify({
                        num: slideNumber,
                        fileUrl: this.fileUrl
                    })
                });
            }, 1000);
        } else if (this.role === Role.Receiver) {
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
        const xhookScript = document.createElement("script");
        xhookScript.src = "//unpkg.com/xhook@latest/dist/xhook.min.js";
        const mutScriptContent = ` 
            xhook.before(request => {
                if (!request.url.startsWith("http")) return; 
                const reqUrl = new URL(request.url);
                if (reqUrl.hostname.includes("excel")) {
                    reqUrl.pathname = "/proxy/" + encodeURIComponent(reqUrl.hostname) + reqUrl.pathname;
                    reqUrl.hostname = "${this.baseUrl.toString().substring(8)}";
                    request.url = reqUrl.toString();
                }
            });

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
        `;
        mutScript.src = "data:text/javascript;charset=utf-8," + escape(mutScriptContent);
        const base = document.createElement("base");
        const url = new URL(this.baseUrl);;
        const proxyComponent = (this.proxyBase === undefined) ? "" : `proxy/${encodeURIComponent(this.proxyBase)}/`;
        base.setAttribute("href", `${url.protocol}//${url.hostname}${url.port ? ":" + url.port : ""}/${proxyComponent}`);
        const scriptTag: HTMLScriptElement = dom.querySelector("#applicationOuterContainer > script:nth-child(6)");
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
                dom.head.insertBefore(xhookScript, dom.head.firstChild);
                document.open();
                document.write(dom.documentElement.innerHTML);
                document.close();
            });
        } else {
            dom.head.insertBefore(base, dom.head.firstChild);
            dom.head.insertBefore(mutScript, dom.head.firstChild);
            dom.head.insertBefore(xhookScript, dom.head.firstChild);
            document.open();
            document.write(dom.documentElement.innerHTML);
            document.close();
        }
    }

    stopSync() {
        this.channel.leave();
    }
}