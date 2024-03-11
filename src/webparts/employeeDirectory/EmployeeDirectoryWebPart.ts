import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-webpart-base';
import styles from './EmployeeDirectoryWebPart.module.scss';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IEmployeeDirectoryWebPartProps {
  description: string;
}

export interface IUser {
  displayName: string;
  jobTitle: string;
  mail: string;
  userPrincipalName: string;
  department: string;
  officeLocation: string;
  mobilePhone: string;
  id: string;
  [key: string]: any;
}

interface GraphApiResponse {
  value: IUser[];
  '@odata.nextLink'?: string;
}

export default class EmployeeDirectoryWebPart extends BaseClientSideWebPart<IEmployeeDirectoryWebPartProps> {
  // private _visibleUsers: IUser[] = [];
  // private _pageSize: number = 10;
  // private _currentPage: number = 1;
  // private _totalPages: number = 0;
  // private _nextPageUrl: string | null = null;
  private _visibleUsers: IUser[] = [];
  private _pageSize: number = 8;
  private _currentPage: number = 1;
  private _nextPageUrl: string | null = null;

  protected onInit(): Promise<void> {
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.1/css/all.min.css');
    return super.onInit();
  }

  public async render(): Promise<void> {
    this.domElement.innerHTML = `
      <div class="${styles.employeeDirectory}">
      <div class="${styles.searchBox}">
        <input type="text" id="searchBox" placeholder="Search users..." />
      </div>
      <div class="${styles.filterBox}">
        <select id="jobTitleFilter">
          <option value="">Loading...</option>
        </select>
        <select id="departmentFilter">
          <option value="">Loading...</option>
        </select>
        <select id="officeLocationFilter">
          <option value="">Loading...</option>
        </select>
        <button id="resetFiltersBtn" class="resetFiltersBtn">Clear Filters</button>
      </div>
      <!--<button id="resetFiltersBtn" class="resetFiltersBtn">Reset Filters</button>-->
      <div class="${styles.userList}" id="userList">
        <div class="${styles.loading}" id="loading"></div>
      </div>
      <div class="${styles.pagination}" id="pagination"></div>
  </div>
  `;
  
    await this.runConcurrently();
  }

  private async runConcurrently(): Promise<void> {
    const populateFilterDropdownsPromise = this._populateFilterDropdowns();
    const setEventHandlersPromise = this._setEventHandlers();
    const getUserDetailsPromise = this._getUserDetails();

    await Promise.all([setEventHandlersPromise, getUserDetailsPromise, populateFilterDropdownsPromise]);

    console.log('All tasks completed.');
  }

  private async _populateFilterDropdowns(): Promise<void> {
    const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');

    const [jobTitles, departments, locations] = await Promise.all([
        this._getUniqueValues(client, 'jobTitle'),
        this._getUniqueValues(client, 'department'),
        this._getUniqueValues(client, 'officeLocation'),
    ]);

    this._populateDropdown('jobTitleFilter', jobTitles);
    this._populateDropdown('departmentFilter', departments);
    this._populateDropdown('officeLocationFilter', locations);

    this._addFilterEventListeners();
  }

  private async _getUniqueValues(client: MSGraphClientV3, field: string): Promise<string[]> {
    const uniqueValues: string[] = [];
    let nextLink: string | null = `/users?$select=${field}&$top=100`; // Start with a small batch size

    while (nextLink) {
        const usersResponse: GraphApiResponse = await client.api(nextLink).version('v1.0').get();
        const users = usersResponse.value;
        users.forEach(user => {
            const value = user[field];
            if (value && !uniqueValues.includes(value)) {
                uniqueValues.push(value);
            }
        });
        nextLink = usersResponse['@odata.nextLink'] || null;
    }

    return uniqueValues;
  }

  private _populateDropdown(dropdownId: string, options: string[]): void {
    const dropdown = document.getElementById(dropdownId) as HTMLSelectElement;
    if (dropdown) {
      dropdown.innerHTML = '';
  
      const defaultOption = document.createElement('option');
      defaultOption.value = '';
      defaultOption.textContent = dropdownId === 'jobTitleFilter' ? '-- Select Job Title --' :
                                  dropdownId === 'departmentFilter' ? '-- Select Department --' :
                                  '-- Select Company --';
      dropdown.appendChild(defaultOption);
  
      options.forEach(option => {
        const optionElement = document.createElement('option');
        optionElement.value = option;
        optionElement.textContent = option;
        dropdown.appendChild(optionElement);
      });
    }
  }

  private async _getUserDetails(): Promise<void> {
    if (!this._nextPageUrl && this._currentPage === 1) {
      this._nextPageUrl = `/users?$select=displayName,jobTitle,mail,companyName,mobilePhone,userPrincipalName,department,officeLocation,id&$top=${this._pageSize}`;
    }

    if (this._nextPageUrl) {
      this._showLoading(true);
      const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      try {
        const usersResponse = await client.api(this._nextPageUrl).version('v1.0').get();
        this._visibleUsers = usersResponse.value;
        this._nextPageUrl = usersResponse['@odata.nextLink'];
        this._renderUserList();
        this._renderPagination();
      } catch (error) {
        console.error("Error fetching users from Microsoft Graph", error);
      } finally {
        this._showLoading(false);
      }
    }
  }

  private _renderUserList(): void {
    let userListHTML: string = '';
    this._visibleUsers.forEach((user: IUser) => {

      const userImageSrc = user.userPrincipalName
        ? `/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(user.userPrincipalName)}`
        : 'url_to_default_image.png'; 
  
        userListHTML += `
        <div class="${styles.userCard}">
          <div class="${styles.userAvatar}">
            <!-- Use SharePoint's _layouts/15/userphoto.aspx for the image source -->
            <img src="${userImageSrc}" onerror="this.onerror=null;this.src='url_to_default_image.png';" alt="User Avatar" />
          </div>
          <div class="${styles.userInfo}">
            <div class="${styles.displayName}">${escape(user.displayName)}</div>
            <div class="${styles.jobTitle}">${escape(user.jobTitle)}</div>
            <div class="${styles.department}">${escape(user.department)}</div>
            <div class="${styles.officeLocation}">${escape(user.officeLocation)}</div>
            <div class="${styles.userPrincipalName}">${escape(user.userPrincipalName)}</div>
            <div class="${styles.mobilePhone}">${escape(user.mobilePhone)}</div>
            <!--<div class="${styles.mail}">${escape(user.mail)}</div>-->
          </div>
          <div class="${styles.contactIcons}">
            <a href="mailto:${escape(user.mail)}" title="Email ${escape(user.displayName)}">
              <i class="fas fa-envelope"></i>
            </a>
            <a href="https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(user.userPrincipalName)}" title="Message on Teams">
            <i class="fa fa-users"></i> 
            </a>
          </div>
        </div>`;
    });
  
    const listContainer = this.domElement.querySelector(`.${styles.userList}`);
    if (listContainer) {
      listContainer.innerHTML = userListHTML;
    }
  }

  private _renderPagination(): void {
    const paginationContainer = this.domElement.querySelector(`.${styles.pagination}`);
    if (paginationContainer) {
      paginationContainer.innerHTML = `
        <button id="prevPageBtn" ${this._currentPage === 1 ? 'disabled' : ''}>Previous</button>
        <span>Page ${this._currentPage}</span>
        <button id="nextPageBtn" ${!this._nextPageUrl ? 'disabled' : ''}>Next</button>
      `;
      this._setPaginationEventHandlers(); // Reattach event handlers
    }
  }

  private _setPaginationEventHandlers(): void {
    const prevPageBtn = this.domElement.querySelector('#prevPageBtn');
    const nextPageBtn = this.domElement.querySelector('#nextPageBtn');
  
    prevPageBtn?.removeEventListener('click', this._prevPageBound);
    nextPageBtn?.removeEventListener('click', this._nextPageBound);
  
    this._prevPageBound = this._prevPage.bind(this);
    this._nextPageBound = this._nextPage.bind(this);
  
    prevPageBtn?.addEventListener('click', this._prevPageBound);
    nextPageBtn?.addEventListener('click', this._nextPageBound);
  }

  private _addFilterEventListeners(): void {
    this.domElement.querySelector('#jobTitleFilter')?.addEventListener('change', (event) => this._filterUsers(event, 'jobTitle'));
    this.domElement.querySelector('#departmentFilter')?.addEventListener('change', (event) => this._filterUsers(event, 'department'));
    this.domElement.querySelector('#officeLocationFilter')?.addEventListener('change', (event) => this._filterUsers(event, 'officeLocation'));
    this.domElement.querySelector('#resetFiltersBtn')?.addEventListener('click', this._resetFilters.bind(this));
  }
  
  private _setEventHandlers(): void {
    const searchBox = this.domElement.querySelector('#searchBox') as HTMLInputElement;
    if (searchBox) {
      searchBox.addEventListener('input', this._searchUsers.bind(this));
    }
  }
  
  private _prevPageBound = this._prevPage.bind(this);
  private _nextPageBound = this._nextPage.bind(this);

  private _pageUrls: string[] = [];

  private _nextPage(): void {
    if (this._nextPageUrl) {

      this._pageUrls[this._currentPage] = this._nextPageUrl;
  
      this._currentPage++;
      this._getUserDetails().catch(error => {
        console.error("Error fetching the next page of users:", error);
        this._currentPage--;
        this._showLoading(false);
      });
    }
  }
  
  private _prevPage(): void {
    if (this._currentPage > 1) {

      this._currentPage--;
  
      this._nextPageUrl = this._pageUrls[this._currentPage - 1];
      this._getUserDetails().catch(error => {
        console.error("Error fetching the previous page of users:", error);
        this._currentPage++;
        this._showLoading(false);
      });
    }
  }  
  
  private async _searchUsers(event: Event): Promise<void> {
    const target = event.target as HTMLInputElement; 
    const searchText = target.value.trim(); 
  
    this._showLoading(true);
    
    if (!searchText) {

      this._currentPage = 1; 
      this._nextPageUrl = null; 
      await this._getUserDetails(); 
      this._showLoading(false); 
      return;
    }
  
    const names = searchText.split(' ').filter(n => n); 
    let filterQuery = names.map(name => `(startsWith(givenName,'${name}') or startsWith(surname,'${name}'))`).join(' and ');
  
    const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
    try {
      let searchUrl = `/users?$filter=${filterQuery}&$select=displayName,jobTitle,mail,companyName,userPrincipalName,department,officeLocation,id&$top=${this._pageSize}`;
      const searchResponse = await client.api(searchUrl).version('v1.0').get();
  
      this._visibleUsers = searchResponse.value;
      this._renderUserList();
      this._renderPagination();
    } catch (error) {
      console.error("Error searching users from Microsoft Graph", error);
    } finally {
      this._showLoading(false); 
    }
  }

  private _currentFilters: { [key: string]: string } = {};

  private async _filterUsers(event: Event, filterType: string): Promise<void> {
    const target = event.target as HTMLSelectElement;
    const filterValue = target.value;
  
    if (filterValue) {
      this._currentFilters[filterType] = filterValue;
    } else {
      delete this._currentFilters[filterType];
    }
  
    this._showLoading(true);
  
    this._currentPage = 1;
    this._nextPageUrl = null;
    
    await this._getUserDetailsFiltered();
  
    this._showLoading(false);
  }
  
  // private async _getUserDetailsFiltered(): Promise<void> {
  //   this._showLoading(true);
  
  //   const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
  //   let filterQueries = [];
  //   for (let [key, value] of Object.entries(this._currentFilters)) {
  //     let encodedValue = encodeURIComponent(value);
  //     filterQueries.push(`${key} eq '${encodedValue}'`);
  //   }
  //   let filterQuery = filterQueries.join(' and ');
  
  //   // Set up the headers required for the advanced query capabilities
  //   const headers = new Headers();
  //   headers.append('ConsistencyLevel', 'eventual'); // Adding the ConsistencyLevel header
  
  //   let apiUrl = `/users?$select=displayName,jobTitle,mail,userPrincipalName,department,officeLocation,id&$top=${this._pageSize}&$count=true`; // Adding $count=true to the query
  //   if (filterQuery) {
  //     apiUrl += `&$filter=${filterQuery}`;
  //   }
  
  //   try {
  //     // Include the headers in the API request
  //     const usersResponse = await client.api(apiUrl).version('v1.0').headers(headers).get();
  //     this._visibleUsers = usersResponse.value;
  //     this._nextPageUrl = usersResponse['@odata.nextLink'] || null;
  //     this._renderUserList();
  //     this._renderPagination();
  //   } catch (error) {
  //     console.error("Error fetching filtered users from Microsoft Graph", error);
  //     console.error(`Failed API URL: ${apiUrl}`);
  //   } finally {
  //     this._showLoading(false);
  //   }
  // }
  
  private async _resetFilters(): Promise<void> {
    this._currentFilters = {};
    this._currentPage = 1;
    this._nextPageUrl = null;
  
    // Reset the dropdown values
    const jobTitleDropdown = this.domElement.querySelector('#jobTitleFilter') as HTMLSelectElement;
    const departmentDropdown = this.domElement.querySelector('#departmentFilter') as HTMLSelectElement;
    const officeLocationDropdown = this.domElement.querySelector('#officeLocationFilter') as HTMLSelectElement;
  
    jobTitleDropdown.value = '';
    departmentDropdown.value = '';
    officeLocationDropdown.value = '';
  
    await this._getUserDetailsFiltered();

  }

  private async _getUserDetailsFiltered(): Promise<void> {
    this._showLoading(true);
    const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
  
    let filterQueries = [];
    for (let [key, value] of Object.entries(this._currentFilters)) {
      let encodedValue = encodeURIComponent(value);
      filterQueries.push(`${key} eq '${encodedValue}'`);
    }
    let filterQuery = filterQueries.join(' and ');
  
    // Update your API URL to include $count=true
    let apiUrl = `/users?$select=displayName,jobTitle,mail,userPrincipalName,department,officeLocation,id&$top=${this._pageSize}&$count=true`;
  
    // Append the filter query if there is one
    if (filterQuery) {
      apiUrl += `&$filter=${filterQuery}`;
    }
  
    try {
      // Make the API call with the additional headers
      const usersResponse = await client.api(apiUrl)
                                        .header('ConsistencyLevel', 'eventual')
                                        .get();
      this._visibleUsers = usersResponse.value;
      this._nextPageUrl = usersResponse['@odata.nextLink'] || null;
      this._renderUserList();
      this._renderPagination();
    } catch (error) {
      console.error("Error fetching filtered users from Microsoft Graph", error);
      console.error(`Failed API URL: ${apiUrl}`);
    } finally {
      this._showLoading(false);
    }
  }
  
  private _showLoading(isLoading: boolean): void {
    const loadingElement = this.domElement.querySelector(`.${styles.loading}`);
    if (loadingElement && loadingElement instanceof HTMLElement) {
      loadingElement.style.display = isLoading ? 'block' : 'none';
    }
  }
  
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Employee Directory Settings"
          },
          groups: [
            {
              groupName: "General Settings",
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Description"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}