import React,{ useEffect } from 'react';
import logo from './logo.svg';
import './App.css';
import { InteractiveBrowserCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import {
  TokenCredentialAuthenticationProvider,
  TokenCredentialAuthenticationProviderOptions,
} from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';

function App() {
  const tokenCredential = new InteractiveBrowserCredential({
    clientId: ""
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
      const user = await client
        .api(`/users/`)
        .get()
        .catch((error: any) => {
          console.log(`Response from Azure AD B2C while getting user: ${error}`);
        });

      console.log(user);
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
