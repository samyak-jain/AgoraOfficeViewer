# AgoraOfficeViewer
A typescript library allowing anyone to quickly render office files in their page with all actions synced across all the users

## Installation:
`npm i agora-office`

## Usage

You can take a look at a sample app built using this library ![here](https://github.com/samyak-jain/AgoraOfficeSample)

## Documentation

### AgoraOffice(backendUrl: string, channel: RtmChannel, iframe: HTMLIFrameElement)

Creates a new AgoraOffice Client Object.  
backendUrl: It is the url pointing to your ![backend](https://github.com/samyak-jain/AgoraOfficeBackend)  
channel: It is an RtmChannel object. You can see how to create that ![here](https://docs.agora.io/en/Real-time-Messaging/messaging_web?platform=Web)  
iframe: Reference to the iframe where you want the document to be rendered  

### setRole(role: Role)

Sets whether the office client should behave as a broacaster or receiver.  
The receiver is always synced to the broadcaster. Example: In a word document the scroll position of the receiver would be in sync with the scroll position of the broadcaster  
*Important: There can be only 1 broadcaster*  

role: The role can be set as a broacaster or receiver.  
- Role.Broadcaster
- Role.Receiver

### loadDocument(fileUrl: URL): Promise<void>

fileUrl: The URL of the document to be loaded into the iframe  
Returns a promise that resolves when the document is done loading  

### syncDocument()

Starts the syncing of the documents.  
If this method is not called for a particular receiver, the syncing will not work only for that receiver.   
If this method is not called for the broadcaster, syncing will not work.  

### stopSync()

Stops the syncing.  
Calling this method for a particular receiver, stops the sync only for that receiver  
Calling this method for the broadcaster, stops the sync for everyone  

### renderFilePickerUi(): Promise<URL>

Renders a file picker UI inside the iframe. It uploads the file to the backend. It is a helper function in case the user does not want to build their own file picker mechanism. Uses ![Dropzone](https://www.dropzonejs.com/) internally.  
The promise resolves after the file is uploaded and returns a URL of the file which can then be passed into the loadDocument() function  





