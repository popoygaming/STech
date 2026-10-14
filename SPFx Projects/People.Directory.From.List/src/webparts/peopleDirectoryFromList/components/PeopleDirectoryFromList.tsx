import * as React from 'react';
import { IconButton, MessageBar, MessageBarType, Persona, PersonaSize, SearchBox, Separator, Spinner, SpinnerSize, Text
} from '@fluentui/react';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import type { IPeopleDirectoryFromListProps } from './IPeopleDirectoryFromListProps';
import { getDefaultSearchTextFromUrl, getFilteredEmployees, getValidPageSize, mapAvailabilityToPresence, toTitleCase } from './PeopleDirectoryFromList.helpers';
import styles from './PeopleDirectoryFromList.module.scss';
import EmployeeDirectoryService, { IEmployeeProfile } from '../services/EmployeeDirectoryService';
import EmployeeProfileService from '../services/EmployeeProfileService';

const LIVE_PERSONA_COMPONENT_ID: string = '914330ee-2df2-4f6e-a858-30c23a812408';

interface ILivePersonaCardLibrary {
  LivePersonaCard: React.ComponentType<React.PropsWithChildren<Record<string, unknown>>>;
}

interface IPeopleDirectoryFromListState {
  employees: IEmployeeProfile[];
  isLoading: boolean;
  currentPage: number;
  searchText: string;
  error?: string;
}

export default class PeopleDirectoryFromList extends React.Component<IPeopleDirectoryFromListProps, IPeopleDirectoryFromListState> {
  private readonly _employeeDirectoryService: EmployeeDirectoryService;
  private readonly _employeeProfileService: EmployeeProfileService;
  private _livePersonaCardLibrary?: ILivePersonaCardLibrary;

  public constructor(props: IPeopleDirectoryFromListProps) {
    super(props);

    this._employeeDirectoryService = new EmployeeDirectoryService(
      this.props.context.spHttpClient,
      this.props.context.pageContext.web.absoluteUrl
    );
    this._employeeProfileService = new EmployeeProfileService(this.props.context);

    this.state = {
      employees: [],
      isLoading: true,
      currentPage: 1,
      searchText: getDefaultSearchTextFromUrl()
    };
  }

  public componentDidMount(): void {
    this._loadLivePersonaCardLibrary().catch((error: unknown) => {
      console.warn('Live Persona Card library could not be loaded.', error);
    });

    this._loadEmployees().catch((error: unknown) => {
      this.setState({
        isLoading: false,
        error: error instanceof Error ? error.message : 'Failed to load employees.'
      });
    });
  }

  public componentDidUpdate(prevProps: IPeopleDirectoryFromListProps): void {
    if (prevProps.employeeListId !== this.props.employeeListId) {
      this.setState({ isLoading: true, error: undefined, currentPage: 1 }, () => {
        this._loadEmployees().catch((error: unknown) => {
          this.setState({
            isLoading: false,
            error: error instanceof Error ? error.message : 'Failed to load employees.'
          });
        });
      });
    }

    if (prevProps.pageSize !== this.props.pageSize) {
      this.setState({ currentPage: 1 });
    }
  }

  public render(): React.ReactElement<IPeopleDirectoryFromListProps> {

    const { employees, isLoading, error, currentPage, searchText } = this.state;
    const filteredEmployees: IEmployeeProfile[] = getFilteredEmployees(employees, searchText);
    const pageSize: number = getValidPageSize(this.props.pageSize);
    const totalPages: number = Math.max(1, Math.ceil(filteredEmployees.length / pageSize));
    const safeCurrentPage: number = Math.min(currentPage, totalPages);
    const startIndex: number = (safeCurrentPage - 1) * pageSize;
    const endIndex: number = Math.min(startIndex + pageSize, filteredEmployees.length);
    const pagedEmployees: IEmployeeProfile[] = filteredEmployees.slice(startIndex, startIndex + pageSize);

    if (isLoading) {
      return (
        <div className={styles.peopleDirectoryFromList}>
          {this.props.showWebPartTitle && <Text variant="xLarge" block={true} className={styles.webPartTitle}>{this.props.webPartTitle}</Text>}
          <Spinner size={SpinnerSize.large} label="Loading employee profiles..." />
        </div>
      );
    }

    if (error) {
      return (
        <div className={styles.peopleDirectoryFromList}>
          {this.props.showWebPartTitle && <Text variant="xLarge" block={true} className={styles.webPartTitle}>{this.props.webPartTitle}</Text>}
          <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
        </div>
      );
    }

    if (employees.length === 0) {
      return (
        <div className={styles.peopleDirectoryFromList}>
          {this.props.showWebPartTitle && <Text variant="xLarge" block={true} className={styles.webPartTitle}>{this.props.webPartTitle}</Text>}
          <MessageBar messageBarType={MessageBarType.info}>No employees found in the Employee Directory list.</MessageBar>
        </div>
      );
    }

    return (
      <div className={styles.peopleDirectoryFromList}>
        {this.props.showWebPartTitle && <Text variant="xLarge" block={true} className={styles.webPartTitle}>{this.props.webPartTitle}</Text>}
        <div className={styles.searchBoxWrapper}>
          <SearchBox
            className={styles.searchBox}
            placeholder="Search by name, email, or job title"
            value={searchText}
            onChange={this._onSearchChange}
          />
          <Separator />
        </div>
        {filteredEmployees.length === 0 && (
          <MessageBar messageBarType={MessageBarType.info}>No matching employees found.</MessageBar>
        )}
        <div className={styles.defaultCard}>
          <div className={styles.documentCardContainer}>
            {pagedEmployees.map((employee: IEmployeeProfile) => {
              const personaElement: JSX.Element = (
                <div className={styles.documentCardItem}>
                  <div className={styles.personaCard}>
                    <Persona
                      text={toTitleCase(employee.fullName)}
                      secondaryText={toTitleCase(employee.role)}
                      tertiaryText={`Phone: ${employee.phone || 'N/A'}`}
                      optionalText={employee.email}
                      size={PersonaSize.size72}
                      imageUrl={employee.profileImageUrl}
                      presence={mapAvailabilityToPresence(employee.availability)}
                      presenceTitle={employee.activity || employee.availability}
                      imageShouldFadeIn={false}
                      hidePersonaDetails={false}
                    />
                  </div>
                </div>
              );

              return (
                <React.Fragment key={employee.id}>
                  {this._renderWithLivePersonaCard(employee, personaElement)}
                </React.Fragment>
              );
            })}
          </div>
        </div>
        {totalPages > 1 && (
          <div className={styles.pagination}>
            <Text variant="small" className={styles.pageSummary}>
              Showing {startIndex + 1}-{endIndex} of {filteredEmployees.length}
            </Text>
            <div className={styles.paginationControls}>
            <IconButton
              iconProps={{ iconName: 'ChevronLeft' }}
              title="Previous page"
              ariaLabel="Previous page"
              className={styles.paginationButton}
              onClick={this._onPreviousPage}
              disabled={safeCurrentPage <= 1}
            />
            <Text variant="small" className={styles.pageInfo}>Page {safeCurrentPage} of {totalPages}</Text>
            <IconButton
              iconProps={{ iconName: 'ChevronRight' }}
              title="Next page"
              ariaLabel="Next page"
              className={styles.paginationButton}
              onClick={this._onNextPage}
              disabled={safeCurrentPage >= totalPages}
            />
            </div>
          </div>
        )}
      </div>
    );
  }

  private async _loadEmployees(): Promise<void> {
    if (!this.props.employeeListId || this.props.employeeListId.trim().length === 0) {
      this.setState({
        employees: [],
        isLoading: false,
        error: 'Please configure Employee List ID in web part properties.'
      });
      return;
    }

    const employees: IEmployeeProfile[] = await this._employeeDirectoryService.getEmployees(this.props.employeeListId);

    const enrichedEmployees: IEmployeeProfile[] = await Promise.all(
      employees.map(async (employee: IEmployeeProfile): Promise<IEmployeeProfile> => {
        const visuals = await this._employeeProfileService.getProfileVisuals(employee.email);

        return {
          ...employee,
          profileImageUrl: visuals.profileImageUrl,
          availability: visuals.availability,
          activity: visuals.activity
        };
      })
    );

    this.setState({
      employees: enrichedEmployees,
      isLoading: false,
      currentPage: 1,
      error: undefined
    });
  }

  private _onPreviousPage = (): void => {
    this.setState((prevState: IPeopleDirectoryFromListState) => ({
      currentPage: Math.max(1, prevState.currentPage - 1)
    }));
  };

  private _onNextPage = (): void => {
    const filteredEmployees: IEmployeeProfile[] = getFilteredEmployees(this.state.employees, this.state.searchText);
    const totalPages: number = Math.max(1, Math.ceil(filteredEmployees.length / getValidPageSize(this.props.pageSize)));

    this.setState((prevState: IPeopleDirectoryFromListState) => ({
      currentPage: Math.min(totalPages, prevState.currentPage + 1)
    }));
  };

  private _onSearchChange = (_event?: React.ChangeEvent<HTMLInputElement>, newValue?: string): void => {
    this.setState({
      searchText: newValue || '',
      currentPage: 1
    });
  };

  private _renderWithLivePersonaCard(employee: IEmployeeProfile, child: JSX.Element): JSX.Element {
    const identifier: string = (employee.email || '').trim();

    if (!this._livePersonaCardLibrary || !identifier) {
      return child;
    }

    const LivePersonaCard: React.ComponentType<React.PropsWithChildren<Record<string, unknown>>> = this._livePersonaCardLibrary.LivePersonaCard;

    return (
      <LivePersonaCard
        className={styles.livePersonaCard}
        clientScenario="PeopleWebPart"
        disableHover={false}
        serviceScope={this.props.context.serviceScope}
        upn={identifier}
        email={identifier}
      >
        {child}
      </LivePersonaCard>
    );
  }

  private async _loadLivePersonaCardLibrary(): Promise<void> {
    if (Environment.type === EnvironmentType.Local || this._livePersonaCardLibrary) {
      return;
    }

    const loadedComponent: unknown = await SPComponentLoader.loadComponentById(LIVE_PERSONA_COMPONENT_ID);
    const livePersonaLibrary: ILivePersonaCardLibrary = loadedComponent as ILivePersonaCardLibrary;

    if (livePersonaLibrary?.LivePersonaCard) {
      this._livePersonaCardLibrary = livePersonaLibrary;
      this.forceUpdate();
    }
  }
}
