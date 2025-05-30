from uplink import Body, Consumer, Path, Query, error_handler, get, headers, params, post


class ApiError(Exception):
    pass


def raise_api_error(exc_type, exc_val, exc_tb):
    raise ApiError(exc_val)


@error_handler(raise_api_error)
class AppdController(Consumer):
    """Minimal python client for the AppDynamics API"""

    jsessionid: str = None
    xcsrftoken: str = None

    def __init__(self, *args, session=None, **kwargs):
        super().__init__(*args, **kwargs)
        self.client_session=session

    def get_client_session(self):
        return self.client_session

    @params({"action": "login"})
    @get("/controller/auth")
    def login(self):
        """Verifies Login Success (Basic Auth)"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/x-www-form-urlencoded"})
    @post("/controller/api/oauth/access_token")
    def loginOAuth(self, data: Body):
        """Method to get a token."""

    @params({"output": "json"})
    @get("/controller/rest/applications")
    def getApmApplications(self):
        """Retrieves Applications"""

    @params({"output": "json"})
    @get("/controller/rest/applications/{applicationID}/business-transactions")
    def getBTs(self, applicationID: Path):
        """Retrieves Applications"""

    @params({"output": "json"})
    @get("/controller/rest/applications/{applicationID}/nodes")
    def getNodes(self, applicationID: Path):
        """Retrieves Nodes"""

    @params({"output": "json"})
    @get("/controller/rest/applications/{applicationID}/nodes/{nodeID}")
    def getNode(self, applicationID: Path, nodeID: Path):
        """Retrieves an Individual Node"""

    @params({"output": "json"})
    @get("/controller/rest/applications/{applicationID}/tiers")
    def getTiers(self, applicationID: Path):
        """Retrieves Tiers"""

    @params({"output": "json"})
    @get("/controller/restui/transactionConfigProto/getRules/{applicationID}")
    def getBtMatchRules(self, applicationID: Path):
        """Retrieves Business Transaction Match Rules"""

    @params({"output": "json"})
    @get("/controller/restui/transactionConfig/getAppLevelBTConfig/{applicationID}")
    def getAppLevelBTConfig(self, applicationID: Path):
        """Retrieves Application Level Business Transaction Configurations"""

    @params({"output": "json"})
    @get("/controller/rest/applications/{applicationID}/backends")
    def getBackends(self, applicationID: Path):
        """Retrieves Backends"""

    @params({"output": "json"})
    @get("/controller/rest/configuration")
    def getConfigurations(self):
        """Retrieves Controller Configurations"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/restui/customExitPoint/getAllCustomExitPoints")
    def getAllCustomExitPoints(self, application: Body):
        """Retrieves Custom Edit Point Configurations"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/restui/backendConfig/getBackendDiscoveryConfigs")
    def getBackendDiscoveryConfigs(self, body: Body):
        """Retrieves Controller Configurations"""

    @params({"output": "json"})
    @get("/controller/restui/applicationManagerUiBean/getDevModeConfig/{applicationID}")
    def getDevModeConfig(self, applicationID: Path):
        """Retrieves Developer Mode Configuration"""

    @params({"output": "json"})
    @get("/controller/restui/applicationManagerUiBean/instrumentationLevel/{applicationID}")
    def getInstrumentationLevel(self, applicationID: Path):
        """Retrieves Instrumentation Level"""

    @params({"output": "json"})
    @headers(
        {
            "Accept": "application/json, text/plain, */*",
        }
    )
    @get("/controller/restui/agentManager/getAllApplicationComponentsWithNodes/{applicationID}")
    def getAllApplicationComponentsWithNodes(self, applicationID: Path):
        """Retrieves Node Configurations"""

    @params({"output": "json"})
    @headers(
        {
            "Accept": "application/json, text/plain, */*",
        }
    )
    @post("/controller/restui/agentManager/getAgentConfiguration")
    def getAgentConfiguration(self, body: Body):
        """Retrieves Agent Configurations"""

    @params({"output": "json"})
    @get("/controller/restui/applicationManagerUiBean/applicationConfiguration/{applicationID}")
    def getApplicationConfiguration(self, applicationID: Path):
        """Retrieves Application Configuration"""

    @params({"output": "json"})
    @get("/controller/restui/components/application/{applicationID}/components")
    def getApplicationComponents(self, applicationID: Path):
        """Retrieves Application Components for Later  to get getServiceEndpointCustomMatchRules"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/restui/serviceEndpoint/getAll")
    def getServiceEndpointCustomMatchRules(self, body: Body):
        """Retrieves Service Endpoint Custom Match Rules for an individual Application Tier"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/restui/serviceEndpoint/getServiceEndpointMatchConfigs")
    def getServiceEndpointDefaultMatchRules(self, body: Body):
        """Retrieves Service Endpoint Custom Match Rules for an individual Application Tier"""

    @params({"output": "json"})
    @get("/controller/restui/events/eventCounts")
    def getEventCounts(
            self,
            applicationID: Query("applicationId"),
            entityType: Query("entityType"),
            entityID: Query("entityId"),
            timeRangeString: Query("timeRangeString"),
    ):
        """Retrieves Event Counts"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/restui/metricBrowser/async/metric-tree/root")
    def getMetricTree(self, body: Body):
        """Retrieves Metrics"""

    @params({"output": "json"})
    @get("/controller/rest/applications/{applicationID}/metric-data")
    def getMetricData(
            self,
            applicationID: Path,
            metric_path: Query("metric-path"),
            rollup: Query("rollup"),
            time_range_type: Query("time-range-type"),
            duration_in_mins: Query("duration-in-mins"),
            start_time: Query("start-time"),
            end_time: Query("end-time"),
    ):
        """Retrieves Metrics"""

    @params({"output": "json"})
    @get("/controller/rest/applications/{applicationID}/events")
    def getApplicationEvents(
            self,
            applicationID: Path,
            event_types: Query("event-types"),
            severities: Query("severities"),
            time_range_type: Query("time-range-type"),
            duration_in_mins: Query("duration-in-mins"),
            start_time: Query("start-time"),
            end_time: Query("end-time"),
    ):
        """Retrieves Events"""

    @params({"output": "json"})
    @get("/controller/alerting/rest/v1/applications/{applicationID}/health-rules")
    def getHealthRules(self, applicationID: Path):
        """Retrieves Health Rules"""

    @params({"output": "json"})
    @get("/controller/alerting/rest/v1/applications/{applicationID}/health-rules/{healthRuleID}")
    def getHealthRule(self, applicationID: Path, healthRuleID: Path):
        """Retrieves Specific Health Rule"""

    @params({"output": "json"})
    @get("/controller/alerting/rest/v1/applications/{applicationID}/policies")
    def getPolicies(self, applicationID: Path):
        """Retrieves Policies"""

    @params({"output": "json"})
    @get("/controller/restui/MidcUiService/getAllDataGathererConfigs/{applicationID}")
    def getDataCollectors(self, applicationID: Path):
        """Retrieves Data Collectors"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/restui/snapshot/snapshotListDataWithFilterHandle")
    def getSnapshotsWithDataCollector(self, body: Body):
        """Retrieves Snapshots"""

    @params({"output": "json"})
    @get("/controller/restui/analyticsConfigTxnAnalyticsUiService/getAllVisibleAppsWithAnalyticsInfo")
    def getAnalyticsEnabledStatusForAllApplications(self):
        """Retrieves Analytics Enabled Status for app Applications"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @get("/controller/restui/dashboards/getAllDashboardsByType/false")
    def getAllDashboardsMetadata(self):
        """Retrieves all Dashboards"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @get("/controller/CustomDashboardImportExportServlet")
    def getDashboard(self, dashboardId: Query("dashboardId")):
        """Retrieves a single Dashboard"""

    @params({"output": "json"})
    @get("/controller/restui/userAdministrationUiService/users")
    def getUsers(self):
        """Retrieves list of Users"""

    @params({"output": "json"})
    @get("/controller/restui/userAdministrationUiService/users/{userID}")
    def getUser(self, userID: Path):
        """Retrieves permission set of a given user"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/restui/licenseRule/getAllLicenseModuleProperties")
    def getAccountUsageSummary(self, body: Body):
        """Retrieves license usage summary"""

    @params({"output": "json"})
    @get("/controller/restui/apiClientAdministrationUiService/apiClients")
    def getApiClients(self):
        """Retrieves list of API Clients"""

    @params({"output": "json"})
    @get("/controller/restui/accountRoleAdministrationUiService/accountRoleSummaries")
    def getRoles(self):
        """Retrieves list of Roles"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/restui/licenseRule/getEumLicenseUsage")
    def getEumLicenseUsage(self, body: Body):
        """Retrieves EUM license usage"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/restui/agents/list/appserver")
    def getAppServerAgents(self, body: Body):
        """Retrieves app server agent summary list"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/restui/agents/list/machine")
    def getMachineAgents(self, body: Body):
        """Retrieves machine agent summary list"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/restui/agents/list/appserver/ids")
    def getAppServerAgentsIds(self, body: Body):
        """Retrieves app server agent summary list"""

    @params({"output": "json"})
    @get("/controller/restui/components/getNodeViewData/{applicationId}/{nodeId}")
    def getAppServerAgentsMetadata(self, applicationId: Path, nodeId: Path):
        """Retrieves app agent metadata"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/restui/agents/list/machine/ids")
    def getMachineAgentsIds(self, body: Body):
        """Retrieves machine agent summary list"""

    @params({"output": "json"})
    @get("/controller/restui/agent/setting/getDBAgents")
    def getDBAgents(self):
        """Retrieves db agent summary list"""

    @params({"output": "json"})
    @get("/controller/restui/analytics/agents/agentsStatus")
    def getAnalyticsAgents(self):
        """Retrieves analytics agent summary list"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/sim/v2/user/machines/keys")
    def getServersKeys(self, body: Body):
        """Retrieves machine agents in bulk"""

    @params({"output": "json"})
    @get("/controller/sim/v2/user/machines/{machineId}")
    def getServer(self, machineId: Path):
        """Retrieves server agent info"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/sim/v2/user/metrics/query/machines")
    def getServerAvailability(self, body: Body):
        """Retrieves server availability info"""

    @params({"output": "json"})
    @get("/controller/restui/eumApplications/getAllEumApplicationsData")
    def getEumApplications(self, timeRange: Query("time-range")):
        """Retrieves all Eum Applications"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/restui/pageList/getEumPageListViewData")
    def getEumPageListViewData(self, body: Body):
        """Retrieves Eum Page List View Data"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/restui/web/pagelist")
    def getEumNetworkRequestList(self, body: Body):
        """Retrieves Eum Network Request List"""

    @params({"output": "json"})
    @get("/controller/restui/browserRUMConfig/getPagesAndFramesConfig/{applicationId}")
    def getPagesAndFramesConfig(self, applicationId: Path):
        """Retrieves pages and frames config"""

    @params({"output": "json"})
    @get("/controller/restui/browserRUMConfig/getAJAXConfig/{applicationId}")
    def getAJAXConfig(self, applicationId: Path):
        """Retrieves AJAX config"""

    @params({"output": "json"})
    @get("/controller/restui/browserRUMConfig/getVirtualPagesConfig/{applicationId}")
    def getVirtualPagesConfig(self, applicationId: Path):
        """Retrieves virtual pages config"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/restui/browserSnapshotList/getSnapshots")
    def getBrowserSnapshots(self, body: Body):
        """Retrieves browser snapshots"""

    @params({"output": "json"})
    @get("/controller/restui/eumApplications/getAllMobileApplicationsData")
    def getMRUMApplications(self, timeRange: Query("time-range")):
        """Retrieves all Mrum Applications"""

    @params({"output": "json"})
    @get("/controller/restui/mobileRUMConfig/networkRequestsConfig/{applicationId}")
    def getMRUMNetworkRequestConfig(self, applicationId: Path):
        """Retrieves Mrum network requests config"""

    @params({"output": "json"})
    @get("/controller/restui/mobileRequestListUiService/getNetworkRequestLimit/{applicationId}")
    def getNetworkRequestLimit(self, applicationId: Path):
        """Retrieves network request limit"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/restui/mobileSnapshotListUiService/getMobileSnapshotSummaries")
    def getMobileSnapshots(self, body: Body):
        """Retrieves mobile snapshots"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/restui/synthetic/schedule/getJobList/{applicationId}")
    def getSyntheticJobs(self, applicationId: Path):
        """Retrieves Synthetic Job List"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/restui/eumSyntheticJobListUiService/getBillableTimeData")
    def getSyntheticBillableTime(self, body: Body):
        """Retrieves Synthetic Billable Time"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/restui/synthetic/schedule/{applicationId}/getJobPAUtilizations")
    def getSyntheticPrivateAgentUtilization(self, applicationId: Path, body: Body):
        """Retrieves Synthetic Private Agent Utilization"""

    @params({"output": "json"})
    @headers({"Content-Type": "application/json"})
    @post("/controller/restui/eumSyntheticJobListUiService/getSessionData")
    def getSyntheticSessionData(self, body: Body):
        """Retrieves Synthetic Session Data"""

    @params({"output": "json"})
    @get("/controller/restui/report/list")
    def getReportList(self):
        """Retrieves Report Data"""
