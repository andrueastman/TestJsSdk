import React,{ useEffect } from 'react';
import logo from './logo.svg';
import './App.css';
import * as fs from "fs";
import { InteractiveBrowserCredential } from '@azure/identity';
import { Client, LargeFileUploadTask, FileUpload, UploadEventHandlers, OneDriveLargeFileUploadOptions, OneDriveLargeFileUploadTask, UploadResult, LargeFileUploadTaskOptions } from '@microsoft/microsoft-graph-client';
import {
  TokenCredentialAuthenticationProvider,
  TokenCredentialAuthenticationProviderOptions,
} from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';

function App() {
  const tokenCredential = new InteractiveBrowserCredential({
    clientId: "d662ac70-7482-45af-9dc3-c3cde8eeede4"
  });

  const options: TokenCredentialAuthenticationProviderOptions = {
    scopes: ['https://graph.microsoft.com/.default'],
  };

  // Create an instance of the TokenCredentialAuthenticationProvider by passing the tokenCredential instance and options to the constructor
  const authProvider = new TokenCredentialAuthenticationProvider(
    tokenCredential,
    options,
  );

  const client = Client.initWithMiddleware({
    debugLogging: true,
    authProvider: authProvider,
  });

  // Similar to componentDidMount and componentDidUpdate:
  useEffect(() => {
    const fetchData = async () => {

      // Make simple request to graph
      const user = await client
        .api(`/users/`)
        .get()
        .catch((error: any) => {
          console.log(`Response from Azure AD B2C while getting user: ${error}`);
        });

      console.log(user);

      const payload = {
        item: {
          "@microsoft.graph.conflictBehavior": "rename",
          name: "<FILE_NAME>",
        },
      };

      // Try to make a upload of a large file
      const uploadSession = await LargeFileUploadTask.createUploadSession(client, "REQUEST_URL", payload); // TODO bug in docs
      const fileName = "<FILE_NAME>";
      const stats = fs.statSync(`./test/sample_files/${fileName}`);
      const totalsize = stats.size;
      const readStream = fs.readFileSync(`./test/sample_files/${fileName}`);
      const fileObject = new FileUpload(readStream, fileName, totalsize);

      const progress = (range?: Range, extraCallBackParam?: unknown) => {
        // Handle progress event
        console.log(range);
      };
      
      const uploadEventHandlers: UploadEventHandlers = {
        progress,
        extraCallBackParam: true,
      };
      
      const options: LargeFileUploadTaskOptions = {
        rangeSize: 327680,
        uploadEventHandlers: uploadEventHandlers,
      };

      const uploadTask = new LargeFileUploadTask(client, fileObject, uploadSession, options);
      const uploadResult: UploadResult = await uploadTask.upload();

      console.log(uploadResult);
    }

    fetchData();
  });

  
  
  return (
    <div className="App">
      <header className="App-header">
        <img src={logo} className="App-logo" alt="logo" />
        <p>
          Edit <code>src/App.tsx</code> and save to reload.
        </p>
        <a
          className="App-link"
          href="https://reactjs.org"
          target="_blank"
          rel="noopener noreferrer"
        >
          Learn React
        </a>
      </header>
    </div>
  );
}

export default App;
