// Adapted from the original source: https://github.com/DirtyHairy/typescript-deferred
// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

module Kurve {


    function DispatchDeferred(closure: () => void) {
        setTimeout(closure, 0);
    }

    enum PromiseState { Pending, ResolutionInProgress, Resolved, Rejected }

    class Client {
        constructor(
            private _dispatcher: (closure: () => void) => void,
            private _successCB: any,
            private _errorCB: any
        ) {
            this.result = new Deferred<any, any>(_dispatcher);
        }

        resolve(value: any, defer: boolean): void {
            if (typeof (this._successCB) !== 'function') {
                this.result.resolve(value);
                return;
            }

            if (defer) {
                this._dispatcher(() => this._dispatchCallback(this._successCB, value));
            } else {
                this._dispatchCallback(this._successCB, value);
            }
        }

        reject(error: any, defer: boolean): void {
            if (typeof (this._errorCB) !== 'function') {
                this.result.reject(error);
                return;
            }

            if (defer) {
                this._dispatcher(() => this._dispatchCallback(this._errorCB, error));
            } else {
                this._dispatchCallback(this._errorCB, error);
            }
        }

        private _dispatchCallback(callback: (arg: any) => any, arg: any): void {
            var result: any,
                then: any,
                type: string;

            try {
                result = callback(arg);
                this.result.resolve(result);
            } catch (err) {
                this.result.reject(err);
                return;
            }
        }

        result: Deferred<any, any>;
    }

    export class Deferred<T, E>  {
        private _dispatcher: (closure: () => void)=> void;

        constructor();
        constructor(dispatcher: (closure: () => void) => void);
        constructor(dispatcher?: (closure: () => void) => void) {
            if (dispatcher)
                this._dispatcher = dispatcher;
            else
                this._dispatcher = DispatchDeferred;
            this.promise = new Promise<T, E>(this);
        }

        private DispatchDeferred(closure: () => void) {
            setTimeout(closure, 0);
        }

        then(successCB: any, errorCB: any): any {
            if (typeof (successCB) !== 'function' && typeof (errorCB) !== 'function') {
                return this.promise;
            }

            var client = new Client(this._dispatcher, successCB, errorCB);

            switch (this._state) {
                case PromiseState.Pending:
                case PromiseState.ResolutionInProgress:
                    this._stack.push(client);
                    break;

                case PromiseState.Resolved:
                    client.resolve(this._value, true);
                    break;

                case PromiseState.Rejected:
                    client.reject(this._error, true);
                    break;
            }

            return client.result.promise;
        }

        resolve(value?: T): Deferred<T, E>;

        resolve(value?: Promise<T, E>): Deferred<T, E>;

        resolve(value?: any): Deferred<T, E> {
            if (this._state !== PromiseState.Pending) {
                return this;
            }

            return this._resolve(value);
        }

        private _resolve(value: any): Deferred<T, E> {
            var type = typeof (value),
                then: any,
                pending = true;

            try {
                if (value !== null &&
                    (type === 'object' || type === 'function') &&
                    typeof (then = value.then) === 'function') {
                    if (value === this.promise) {
                        throw new TypeError('recursive resolution');
                    }

                    this._state = PromiseState.ResolutionInProgress;
                    then.call(value,
                        (result: any): void => {
                            if (pending) {
                                pending = false;
                                this._resolve(result);
                            }
                        },
                        (error: any): void => {
                            if (pending) {
                                pending = false;
                                this._reject(error);
                            }
                        }
                    );
                } else {
                    this._state = PromiseState.ResolutionInProgress;

                    this._dispatcher(() => {
                        this._state = PromiseState.Resolved;
                        this._value = value;

                        var i: number,
                            stackSize = this._stack.length;

                        for (i = 0; i < stackSize; i++) {
                            this._stack[i].resolve(value, false);
                        }

                        this._stack.splice(0, stackSize);
                    });
                }
            } catch (err) {
                if (pending) {
                    this._reject(err);
                }
            }

            return this;
        }

        reject(error?: E): Deferred<T, E> {
            if (this._state !== PromiseState.Pending) {
                return this;
            }

            return this._reject(error);
        }

        private _reject(error?: any): Deferred<T, E> {
            this._state = PromiseState.ResolutionInProgress;

            this._dispatcher(() => {
                this._state = PromiseState.Rejected;
                this._error = error;

                var stackSize = this._stack.length,
                    i = 0;

                for (i = 0; i < stackSize; i++) {
                    this._stack[i].reject(error, false);
                }

                this._stack.splice(0, stackSize);
            });

            return this;
        }

        promise: Promise<T, E>;

        private _stack: Array<Client> = [];
        private _state = PromiseState.Pending;
        private _value: T;
        private _error: any;
    }

    export class Promise<T, E> implements Promise<T, E> {
        constructor(private _deferred: Deferred<T, E>) { }

        then<R>(
            successCallback?: (result: T) => R,
            errorCallback?: (error: E) => R
        ): Promise<R, E>;

        then(successCallback: any, errorCallback: any): any {
            return this._deferred.then(successCallback, errorCallback);
        }

        fail<R>(
            errorCallback?: (error: E) => R
        ): Promise<R, E>;

        fail(errorCallback: any): any {
            return this._deferred.then(undefined, errorCallback);
        }
    }
}

//*********************************************************   
//   
//Kurve js, https://github.com/microsoftdx/kurvejs
//  
//Copyright (c) Microsoft Corporation  
//All rights reserved.   
//  
// MIT License:  
// Permission is hereby granted, free of charge, to any person obtaining  
// a copy of this software and associated documentation files (the  
// ""Software""), to deal in the Software without restriction, including  
// without limitation the rights to use, copy, modify, merge, publish,  
// distribute, sublicense, and/or sell copies of the Software, and to  
// permit persons to whom the Software is furnished to do so, subject to  
// the following conditions:  




// The above copyright notice and this permission notice shall be  
// included in all copies or substantial portions of the Software.  




// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,  
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF  
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND  
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE  
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION  
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION  
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.  
//   
//*********************************************************   

// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
module Kurve {
    export class Error {
        public status: number;
        public statusText: string;
        public text: string;
        public other: any;
    }

     export class Identity {
        public authContext: any = null;
        public config: any = null;
        public isCallback: boolean = false;
        public clientId: string;
        private req: XMLHttpRequest;
        private state: string;
        private nonce: string;
        private idToken: any;
        private loginCallback: (error: Error) => void;
        private accessTokenCallback: (token:string, error: Error) => void;
        private getTokenCallback: (token: string, error: Error) => void;
        private redirectUri: string;
        private tokenCache: any;
        private logonUser: any;
        private refreshTimer: any;
    
        constructor(clientId = "", redirectUri = "") {
            this.clientId = clientId;
            this.redirectUri = redirectUri;
            this.req = new XMLHttpRequest();
            this.tokenCache = {};

            //Callback handler from other windows
            window.addEventListener("message", ((event: MessageEvent) => {
                if (event.data.type === "id_token") {
                    //Callback being called by the login window
                    if (!event.data.token) {
                        this.loginCallback(event.data);
                    }
                    else {
                        //check for state
                        if (this.state !== event.data.state) {
                            var error = new Error();
                            error.statusText = "Invalid state";
                            this.loginCallback(error);
                        } else {
                            this.decodeIdToken(event.data.token);
                            this.loginCallback(null);
                        }
                    }
                } else if (event.data.type === "access_token") {
                    //Callback being called by the iframe with the token
                    if (!event.data.token)
                        this.getTokenCallback(null, event.data);
                    else {
                        var token:string = event.data.token;
                        var iframe = document.getElementById("tokenIFrame");
                        iframe.parentNode.removeChild(iframe);

                        if (event.data.state !== this.state) {
                            var error = new Error();
                            error.statusText = "Invalid state";
                            this.getTokenCallback(null, error);
                        }
                        else {
                            this.getTokenCallback(token, null);
                        }
                    }
                }
            }));
        }

        private decodeIdToken(idToken: string): void {
           
            var decodedToken = this.base64Decode(idToken.substring(idToken.indexOf('.') + 1, idToken.lastIndexOf('.')));
            var decodedTokenJSON = JSON.parse(decodedToken);
            var expiryDate = new Date(new Date('01/01/1970 0:0 UTC').getTime() + parseInt(decodedTokenJSON.exp) * 1000);
            this.idToken = {
                token: idToken,
                expiry: expiryDate,
                upn: decodedTokenJSON.upn,
                tenantId: decodedTokenJSON.tid,
                family_name: decodedTokenJSON.family_name,
                given_name: decodedTokenJSON.given_name,
                name: decodedTokenJSON.name
            }
            var expiration: Number = expiryDate.getTime() - new Date().getTime() - 300000;

            this.refreshTimer = setTimeout((() => {
                this.renewIdToken();
            }), expiration); 
        }

        private decodeAccessToken(accessToken: string, resource:string): void {
            var decodedToken = this.base64Decode(accessToken.substring(accessToken.indexOf('.') + 1, accessToken.lastIndexOf('.')));
            var decodedTokenJSON = JSON.parse(decodedToken);
            var expiryDate = new Date(new Date('01/01/1970 0:0 UTC').getTime() + parseInt(decodedTokenJSON.exp) * 1000);
            this.tokenCache[resource] = {
                resource: resource,
                token: accessToken,
                expiry: expiryDate
            }
        }

        public getIdToken(): any {
            return this.idToken;
        }
        public isLoggedIn(): boolean {
            if (!this.idToken) return false;
            return (this.idToken.expiry > new Date());
        }

        private renewIdToken(): void {
            clearTimeout(this.refreshTimer);
            this.login((() => {

            }));
        }

        public getAccessTokenAsync(resource: string): Promise<string,Error> {

            var d = new Deferred<string,Error>();
            this.getAccessToken(resource, ((token, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(token);
                }
            }));
            return d.promise;
        }

      

        public getAccessToken(resource: string, callback: (token: string, error: Error) => void): void {
            //Check for cache and see if we have a valid token
            var cachedToken = this.tokenCache[resource];
            if (cachedToken) {
                //We have it cached, has it expired? (5 minutes error margin)
                if (cachedToken.expiry > (new Date(new Date().getTime() + 60000))) {
                    callback(<string>cachedToken.token, null);
                    return;
                }
            }
            //If we got this far, we need to go get this token

            //Need to create the iFrame to invoke the acquire token
            this.getTokenCallback = ((token: string, error: Error) => {
                if (error) {
                    callback(null, error);
                }
                else {
                    this.decodeAccessToken(token, resource);
                    callback(token, null);

                }
            });

            this.nonce = this.generateNonce();
            this.state = this.generateNonce();

            var iframe = document.createElement('iframe');
            iframe.style.display = "none";
            iframe.id = "tokenIFrame";
            iframe.src = "./login.html?clientId=" + encodeURIComponent(this.clientId) +
            "&resource=" + encodeURIComponent(resource) +
            "&redirectUri=" + encodeURIComponent(this.redirectUri) +
            "&state=" + encodeURIComponent(this.state) +
            "&nonce=" + encodeURIComponent(this.nonce);
            document.body.appendChild(iframe);
        }

        public loginAsync(): Promise<void, Error> {
            var d = new Deferred<void,Error>();
            this.login(((error) => {
                if (error) {
                    d.reject(error);
                }
                else {
                    d.resolve(null);
                }
            }));
            return d.promise;
        }

        public login(callback: (error: Error) => void): void {
            this.loginCallback = callback;
            this.state = this.generateNonce();
            this.nonce = this.generateNonce();
            window.open("./login.html?clientId=" + encodeURIComponent(this.clientId) +
                "&redirectUri=" + encodeURIComponent(this.redirectUri) +
                "&state=" + encodeURIComponent(this.state) +
                "&nonce=" + encodeURIComponent(this.nonce), "_blank");
        }

        public logOut(): void {
            var url = "https://login.microsoftonline.com/common/oauth2/logout?post_logout_redirect_uri=" + encodeURI(window.location.href);
            window.location.href = url;
        }

        private base64Decode(encodedString: string): string {
            var e = {}, i, b = 0, c, x, l = 0, a, r = '', w = String.fromCharCode, L = encodedString.length;
            var A = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
            for (i = 0; i < 64; i++) { e[A.charAt(i)] = i; }
            for (x = 0; x < L; x++) {
                c = e[encodedString.charAt(x)]; b = (b << 6) + c; l += 6;
                while (l >= 8) { ((a = (b >>> (l -= 8)) & 0xff) || (x < (L - 2))) && (r += w(a)); }
            }
            return r;
        }

        private generateNonce(): string {
            var text = "";
            var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            for (var i = 0; i < 32; i++) {
                text += chars.charAt(Math.floor(Math.random() * chars.length));
            }
            return text;
        }
    }
}


//*********************************************************   
//   
//Kurve js, https://github.com/microsoftdx/kurvejs
//  
//Copyright (c) Microsoft Corporation  
//All rights reserved.   
//  
// MIT License:  
// Permission is hereby granted, free of charge, to any person obtaining  
// a copy of this software and associated documentation files (the  
// ""Software""), to deal in the Software without restriction, including  
// without limitation the rights to use, copy, modify, merge, publish,  
// distribute, sublicense, and/or sell copies of the Software, and to  
// permit persons to whom the Software is furnished to do so, subject to  
// the following conditions:  




// The above copyright notice and this permission notice shall be  
// included in all copies or substantial portions of the Software.  




// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,  
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF  
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND  
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE  
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION  
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION  
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.  
//   
//*********************************************************   

// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
module Kurve {

    export class ProfilePhotoDataModel {
        public id: string;
        public height: Number;
        public width: Number;        
    }

    export class ProfilePhoto {
        constructor(protected graph: Kurve.Graph, protected _data: ProfilePhotoDataModel) {
        }

        get data() { return this._data; }
    }

    export class UserDataModel {
         public businessPhones : string;
         public displayName: string;
         public givenName: string;
         public jobTitle: string;
         public mail: string;
         public mobilePhone: string;
         public officeLocation: string;
         public preferredLanguage: string;
         public surname: string;
         public userPrincipalName: string;
         public id: string;
    }

    export class User  {        
        constructor(protected graph: Kurve.Graph, protected _data: UserDataModel) {
        }

        get data() { return this._data; }

        // These are all passthroughs to the graph

        public memberOf(callback: (groups: Groups, Error) => void, Error, odataQuery?: string) {
            this.graph.memberOfForUser(this._data.userPrincipalName, callback, odataQuery);
        }

        public memberOfAsync(odataQuery?: string): Promise<Messages,Error> {
            return this.graph.memberOfForUserAsync(this._data.userPrincipalName, odataQuery);
        }

        public messages(callback: (messages: Messages, error: Error) => void, odataQuery?: string) {
            this.graph.messagesForUser(this._data.userPrincipalName, callback, odataQuery);
        }

        public messagesAsync(odataQuery?: string): Promise<Messages,Error> {
            return this.graph.messagesForUserAsync(this._data.userPrincipalName, odataQuery);
        }

        public manager(callback: (user: Kurve.User, error: Error) => void, odataQuery?: string) {
            this.graph.managerForUser(this._data.userPrincipalName, callback, odataQuery);
        }

        public managerAsync(odataQuery?: string): Promise<User,Error> {
            return this.graph.managerForUserAsync(this._data.userPrincipalName, odataQuery);
        }      

        public profilePhoto(callback: (photo: ProfilePhoto, error: Error) => void) {
            this.graph.profilePhotoForUser(this._data.userPrincipalName, callback);
        }

        public profilePhotoAsync(): Promise<ProfilePhoto, Error> {
            return this.graph.profilePhotoForUserAsync(this._data.userPrincipalName);
        }

        public profilePhotoValue(callback: (val: any, error: Error) => void) {
            this.graph.profilePhotoValueForUser(this._data.userPrincipalName, callback);
        }

        public profilePhotoValueAsync(): Promise<any, Error> {
            return this.graph.profilePhotoValueForUserAsync(this._data.userPrincipalName);
        }

        public calendar(callback: (calendarItems: CalendarEvents, error: Error) => void, odataQuery?: string) {
            this.graph.calendarForUser(this._data.userPrincipalName, callback, odataQuery);
        }

        public calendarAsync(odataQuery?: string): Promise<CalendarEvents, Error> {
            return this.graph.calendarForUserAsync(this._data.userPrincipalName, odataQuery);
        }

    }

    export class Users {

        public nextLink: (callback?: (users: Kurve.Users, error: Error) => void, odataQuery?: string) => Promise<Users, Error>
        constructor(protected graph: Kurve.Graph, protected _data: User[]) {
        }

        get data(): User[] {
            return this._data;
        }
    }
    export class MessageDataModel {
        bccRecipients: string[]
        body: Object
        bodyPreview: string;
        categories: string[]
        ccRecipients: string[]
        changeKey: string;
        conversationId: string;
        createdDateTime: string;
        from: any;
        graph: any;
        hasAttachments: boolean;
        id: string;
        importance: string;
        isDeliveryReceiptRequested: boolean;
        isDraft: boolean;
        isRead: boolean;
        isReadReceiptRequested: boolean;
        lastModifiedDateTime: string;
        parentFolderId: string;
        receivedDateTime: string;
        replyTo: any[]
        sender: any;
        sentDateTime: string;
        subject: string;
        toRecipients: string[]
        webLink: string;
    }

    export class Message {
        constructor(protected graph: Kurve.Graph, protected _data: MessageDataModel) {
        }
        get data() : MessageDataModel {
            return this._data;
        }
    }

    export class Messages {

        public nextLink: (callback?: (messages: Kurve.Messages, error: Error) => void, odataQuery?: string) => Promise<Messages, Error>
        constructor(protected graph: Kurve.Graph, protected _data: Message[]) {
        }

        get data(): Message[] {
            return this._data;
        }
    }

    export class CalendarEvent {
    }
    export class CalendarEvents {

        public nextLink: (callback?: (events: Kurve.CalendarEvents, error: Error) => void, odataQuery?: string) => Promise<(events: Kurve.CalendarEvents, error: Error) => void,Error>
        constructor(protected graph: Kurve.Graph, protected _data: CalendarEvent[]) {
        }

        get data(): CalendarEvent[] {
            return this._data;
        }
    }

    export class Contact {
    }

    export class GroupDataModel {
        public id: string;
        public description: string;
        public displayName: string;
        public groupTypes: string[];
        public mail: string;
        public mailEnabled: Boolean;
        public mailNickname: string;
        public onPremisesLastSyncDateTime: Date;
        public onPremisesSecurityIdentifier: string;
        public onPremisesSyncEnabled: Boolean;
        public proxyAddresses: string[];
        public securityEnabled: Boolean;
        public visibility: string;      
    }

    export class Group {
        constructor(protected graph: Kurve.Graph, protected _data: GroupDataModel) {
        }

        get data() { return this._data; }

    }

    export class Groups {

        public nextLink: (callback?: (groups: Kurve.Groups, error: Error) => void, odataQuery?: string) => Promise<Groups, Error>
        constructor(protected graph: Kurve.Graph, protected _data: Group[]) {
        }

        get data(): Group[] {
            return this._data;
        }
    }

    export class Graph {
        private req: XMLHttpRequest = null;
        private state: string = null;
        private nonce: string = null;
        private accessToken: string = null;
        private KurveIdentity: Identity = null;
        private defaultResourceID: string = "https://graph.microsoft.com";
        private baseUrl: string = "https://graph.microsoft.com/v1.0/";

        constructor(identityInfo: { identity: Identity });
        constructor(identityInfo: { defaultAccessToken: string });
        constructor(identityInfo: any) {
            if (identityInfo.defaultAccessToken) {
                this.accessToken = identityInfo.defaultAccessToken;
            } else {
                this.KurveIdentity = identityInfo.identity;
            }
        }
      
        //Users
        public meAsync(odataQuery?: string): Promise<User, Error> {
            var d = new Deferred<User,Error>();
            this.me((user, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(user);
                }
            }, odataQuery);
            return d.promise;
        }

        public me(callback: (user: User, error: Error) => void, odataQuery?: string): void {
            var urlString: string = this.buildMeUrl() + "/";
            if (odataQuery) {
                urlString += "?" + odataQuery;
            }
            this.getUser(urlString, callback);
        }

        public userAsync(userId: string): Promise<User, Error> {
            var d = new Deferred<User,Error>();
            this.user(userId, (user, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(user);
                }
            });
            return d.promise;
        }

        public user(userId: string, callback: (user: Kurve.User, error: Error) => void): void {
            var urlString: string = this.buildUsersUrl() + "/" + userId;
            this.getUser(urlString, callback);
        }

        public usersAsync(odataQuery?: string): Promise<Users, Error> {
            var d = new Deferred<Users,Error>();
            this.users((users, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(users);
                }
            }, odataQuery);
            return d.promise;
        }

        public users(callback: (users: Kurve.Users, error: Error) => void, odataQuery?: string): void {
            var urlString: string = this.buildUsersUrl() + "/";
            if (odataQuery) {
                urlString += "?" + odataQuery;
            }
            this.getUsers(urlString, callback);
        }

        //Groups

        public groupAsync(groupId: string): Promise<Group,Error> {
            var d = new Deferred<Group,Error>();
            this.group(groupId, (group, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(group);
                }
            });
            return d.promise;
        }

        public group(groupId: string, callback: (group: any, error: Error) => void): void {
            var urlString: string = this.buildGroupsUrl() + "/" + groupId;
            this.getGroup(urlString, callback);
        }

        public groups(callback: (groups: any, error: Error) => void, odataQuery?: string): void {
            var urlString: string = this.buildGroupsUrl() + "/";
            if (odataQuery) {
                urlString += "?" + odataQuery;
            }
            this.getGroups(urlString, callback);
        }

        public groupsAsync(odataQuery?: string): Promise<Groups,Error> {
            var d = new Deferred<Groups,Error>();
            
            this.groups((groups, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(groups);
                }
            }, odataQuery);
            return d.promise;
        }
        

        // Messages For User
            
        public messagesForUser(userPrincipalName: string, callback: (messages: Messages, error: Error) => void, odataQuery?: string): void {
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/messages";
            if (odataQuery) urlString += "?" + odataQuery;

            this.getMessages(urlString, (result, error) => {
                callback(result, error);
            }, odataQuery);
        }

        public messagesForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<Messages, Error> {
            var d = new Deferred<Messages,Error>();
            this.messagesForUser(userPrincipalName, (messages, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(messages);
                }
            }, odataQuery);
            return d.promise;
        }

        // Calendar For User

        public calendarForUser(userPrincipalName: string, callback: (events: CalendarEvent, error: Error) => void, odataQuery?: string): void {
        // // To BE IMPLEMENTED
        //    var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/calendar/events";
        //    if (odataQuery) urlString += "?" + odataQuery;

        //    this.getMessages(urlString, (result, error) => {
        //        callback(result, error);
        //    }, odataQuery);
        }

        public calendarForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<CalendarEvents, Error> {
            var d = new Deferred<CalendarEvents,Error>();
            // // To BE IMPLEMENTED
            //    this.calendarForUser(userPrincipalName, (events, error) => {
            //        if (error) {
            //            d.reject(error);
            //        } else {
            //            d.resolve(events);
            //        }
            //    }, odataQuery);
            return d.promise;
        }

        // Groups/Relationships For User
        public memberOfForUser(userPrincipalName: string, callback: (groups: Kurve.Groups, error: Error) => void, odataQuery?: string) {
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/memberOf";
            if (odataQuery) urlString += "?" + odataQuery;
            this.getGroups(urlString, callback, odataQuery);
        }

        public memberOfForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<Messages, Error> {
            var d:any = new Deferred<Messages,Error>();
            this.memberOfForUser(userPrincipalName, (result, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(result);
                }
            }, odataQuery);
            return d.promise;
        }

        public managerForUser(userPrincipalName: string, callback: (manager: Kurve.User, error: Error) => void, odataQuery?: string) {
            // need odataQuery;
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/manager";
            this.getUser(urlString, callback);
        }

        public managerForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<User, Error> {
            var d = new Deferred<User,Error>();
            this.managerForUser(userPrincipalName, (result, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(result);
                }
            }, odataQuery);
            return d.promise;
        }

        public directReportsForUser(userPrincipalName: string, callback: (users: Kurve.Users, error: Error) => void, odataQuery?: string) {
            // Need odata query
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/directReports";
            this.getUsers(urlString, callback);
        }

        public directReportsForUserAsync(userPrincipalName: string, odataQuery?: string): Promise<Users, Error> {
            var d = new Deferred<Users,Error>();
            this.directReportsForUser(userPrincipalName, (result, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(result);
                }
            }, odataQuery);
            return d.promise;
        }

        public profilePhotoForUser(userPrincipalName: string, callback: (photo: ProfilePhoto, error: Error) => void) {
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/photo";
            this.getPhoto(urlString, callback);
        }

        public profilePhotoForUserAsync(userPrincipalName: string): Promise<ProfilePhoto, Error> {
            var d = new Deferred<ProfilePhoto,Error>();
            
            this.profilePhotoForUser(userPrincipalName, (result, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(result);
                }
            });
            return d.promise;
        }

        public profilePhotoValueForUser(userPrincipalName: string, callback: (photo: any, error: Error) => void) {
            var urlString = this.buildUsersUrl() + "/" + userPrincipalName + "/photo/$value";
            this.getPhotoValue(urlString, callback);
        }

        public profilePhotoValueForUserAsync(userPrincipalName: string): Promise<any, Error>{
            var d = new Deferred<any,Error>();
            this.profilePhotoValueForUser(userPrincipalName, (result, error) => {
                if (error) {
                    d.reject(error);
                } else {
                    d.resolve(result);
                }
            });
            return d.promise;
        }
    
        //http verbs
        public getAsync(url: string): Promise<string, Error> {
            var d = new Deferred<string,Error>();
            this.get(url, (response, error) => {
                if (!error) {
                    d.resolve(response);
                }
                else {
                    d.reject(error);
                }
            });
            return d.promise;
        }

        public get(url: string, callback: (response: string, error: Error) => void, responseType?: string): void {
            var xhr = new XMLHttpRequest();
            if (responseType)
                xhr.responseType = responseType;
            xhr.onreadystatechange = (() => {
                if (xhr.readyState === 4 && xhr.status === 200) {
                    if (!responseType)
                        callback(xhr.responseText, null);
                    else
                        callback(xhr.response, null);
                } else if (xhr.readyState === 4 && xhr.status !== 200) {
                    callback(null, this.generateError(xhr));
                }
            });

            xhr.open("GET", url);
            this.addAccessTokenAndSend(xhr, (addTokenError: Error) => {
                if (addTokenError) {
                    callback(null, addTokenError);
                }
            });
        }

        private generateError(xhr: XMLHttpRequest): Error {
            var response = new Error();
            response.status = xhr.status;
            response.statusText = xhr.statusText;
            if (xhr.responseType === '' || xhr.responseType === 'text')
                response.text = xhr.responseText;
            else
                response.other = xhr.response;
            return response;

        }

        //Private methods

        private getUsers(urlString, callback: (users: Kurve.Users, error: Error) => void): void {
            this.get(urlString, ((result: string, errorGet: Error) => {
                
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }

                var usersODATA = JSON.parse(result);
                if (usersODATA.error) {
                    var errorODATA = new Error();
                    errorODATA.other = usersODATA.error;
                    callback(null, errorODATA);
                    return;
                }

                var resultsArray = (usersODATA.value ? usersODATA.value : [usersODATA]) as any[];
                var users = new Kurve.Users(this, resultsArray.map(o => {
                    return new User(this, o);
                }));

                //implement nextLink
                var nextLink = usersODATA['@odata.nextLink'];

                if (nextLink) {
                    users.nextLink = ((callback?: (result: Users, error: Error) => void) => {
                        var d = new Deferred<Users,Error>();
                        this.getUsers(nextLink, ((result, error) => {
                            if (callback)
                                callback(result, error);
                            else if (error) {
                                d.reject(error);
                            }
                            else {
                                d.resolve(result);
                            }
                        }));
                        return d.promise;
                    });
                }

                callback(users, null);
            }));
        }

        private getUser(urlString, callback: (user: User, error: Error) => void): void {
            this.get(urlString, (result: string, errorGet: Error) => {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                var userODATA = JSON.parse(result) ;
                if (userODATA.error) {
                    var errorODATA = new Error();
                    errorODATA.other = userODATA.error;
                    callback(null, errorODATA);
                    return;
                }

                var user = new User(this, userODATA);
                callback(user, null);
            });

        }

        private addAccessTokenAndSend(xhr: XMLHttpRequest, callback: (error: Error) => void): void {
            if (this.accessToken) {
                //Using default access token
                xhr.setRequestHeader('Authorization', 'Bearer ' + this.accessToken);
                xhr.send();
            } else {
                //Using the integrated Identity object
                this.KurveIdentity.getAccessToken(this.defaultResourceID, ((token: string, error: Error) => {
                    //cache the token
                    
                    if (error)
                        callback(error);
                    else {
                        xhr.setRequestHeader('Authorization', 'Bearer ' + token);
                        xhr.send();
                        callback(null);
                    }
                }));
            }
        }

        private getMessages(urlString: string, callback: (messages: Messages, error: Error) => void, odataQuery?: string): void {

            var url = urlString;
            if (odataQuery) urlString += "?" + odataQuery;
            this.get(url, ((result: string, errorGet: Error) => {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }

                var messagesODATA = JSON.parse(result);
                if (messagesODATA.error) {
                    var errorODATA = new Error();
                    errorODATA.other = messagesODATA.error;
                    callback(null, errorODATA);
                    return;
                }

                var resultsArray = (messagesODATA.value ? messagesODATA.value : [messagesODATA]) as any[];
                var messages = new Kurve.Messages(this, resultsArray.map(o => {
                    return new Message(this, o);
                }));
                if (messagesODATA['@odata.nextLink']) {
                    messages.nextLink = (callback?: (messages: Messages, error: Error) => void, odataQuery?: string) => {
                        var d = new Deferred<Messages,Error>();
                        
                        this.getMessages(messagesODATA['@odata.nextLink'], (messages, error) => {
                            if (callback)
                                callback(messages, error);
                            else if (error) {
                                d.reject(error);
                            }
                            else {
                                d.resolve(messages);
                            }
                        }, odataQuery);
                        return d.promise;

                    };
                }
                callback(messages,  null);
            }));
        }

        private getGroups(urlString: string, callback: (groups: Kurve.Groups, error: Error) => void, odataQuery?: string): void {

            var url = urlString;
            if (odataQuery) urlString += "?" + odataQuery;
            this.get(url, ((result: string, errorGet: Error) => {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                var groupsODATA = JSON.parse(result);
                if (groupsODATA.error) {
                    var errorODATA = new Error();
                    errorODATA.other = groupsODATA.error;
                    callback(null, errorODATA);
                    return;
                }

                var resultsArray = (groupsODATA.value ? groupsODATA.value : [groupsODATA]) as any[];
                var groups = new Kurve.Groups(this, resultsArray.map(o => {
                    return new Group(this, o);
                }));

                var nextLink = groupsODATA['@odata.nextLink'];

                //implement nextLink
                if (nextLink) {
                    groups.nextLink = ((callback?: (result: Groups, error: Error) => void) => {
                        var d = new Deferred<Groups,Error>();
                        this.getGroups(nextLink, ((result, error) => {
                            if (callback)
                                callback(result, error);
                            else if (error) {
                                d.reject(error);
                            }
                            else {
                                d.resolve(result);
                            }
                        }));
                        return d.promise;
                    });
                }

                callback(groups, null);
            }));
        }

        private getGroup(urlString, callback: (group: Kurve.Group, error: Error) => void): void {
            this.get(urlString, (result: string, errorGet: Error) => {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                var ODATA = JSON.parse(result);
                if (ODATA.error) {
                    var ODATAError = new Error();
                    ODATAError.other = ODATA.error;
                    callback(null, ODATAError);
                    return;
                }
                var group = new Kurve.Group(this, ODATA);

                callback(group, null);
            });

        }

        private getPhoto(urlString, callback: (photo: ProfilePhoto, error: Error) => void): void {
            this.get(urlString, (result: string, errorGet: Error) => {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                var ODATA = JSON.parse(result);
                if (ODATA.error) {
                    var errorODATA = new Error();
                    errorODATA.other = ODATA.error;
                    callback(null, errorODATA);
                    return;
                }
                var photo = new ProfilePhoto(this, ODATA);

                callback(photo, null);
            });
        }

        private getPhotoValue(urlString, callback: (photo: any, error: Error) => void): void {
            this.get(urlString, (result: any, errorGet: Error) => {
                if (errorGet) {
                    callback(null, errorGet);
                    return;
                }
                callback(result, null);
            }, "blob");
        }
        private buildMeUrl(): string {
            return this.baseUrl + "me";
        }
        private buildUsersUrl(): string {
            return this.baseUrl + "/users";
        }
        private buildGroupsUrl(): string {
            return this.baseUrl + "/groups";
        }
    }
}

//*********************************************************   
//   
//Kurve js, https://github.com/microsoftdx/kurvejs
//  
//Copyright (c) Microsoft Corporation  
//All rights reserved.   
//  
// MIT License:  
// Permission is hereby granted, free of charge, to any person obtaining  
// a copy of this software and associated documentation files (the  
// ""Software""), to deal in the Software without restriction, including  
// without limitation the rights to use, copy, modify, merge, publish,  
// distribute, sublicense, and/or sell copies of the Software, and to  
// permit persons to whom the Software is furnished to do so, subject to  
// the following conditions:  




// The above copyright notice and this permission notice shall be  
// included in all copies or substantial portions of the Software.  




// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,  
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF  
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND  
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE  
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION  
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION  
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.  
//   
//*********************************************************   
