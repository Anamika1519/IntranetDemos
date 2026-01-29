import * as React from 'react';
import styles from './Dashboard.module.scss';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { IContextualMenuProps } from '@fluentui/react/lib/ContextualMenu';
// import  {useHistory}  from 'react-router-dom';
import Provider from "../../../GlobalContext/provider";
import {IMyProjectProps} from './IMyProjectProps';
import VerticalSideBar from "../../verticalSideBar/components/VerticalSideBar";
import CustomBreadcrumb from '../../../CustomJSComponents/CustomBreadcrumb/CustomBreadcrumb';
import UserContext from "../../../GlobalContext/context";
import "../../../Assets/Figtree/Figtree-VariableFont_wght.ttf";
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
import HorizontalNavbar from "../../horizontalNavBar/components/HorizontalNavBar";
import { SPFI } from "@pnp/sp/presets/all";
import NewRequest from './NewRequest';
// import { fetchUserInformationList } from "../../../APISearvice/GroupTeamService";

// import Multiselect from "multiselect-react-dropdown";

import { getSP } from "../loc/pnpjsConfig";
import { useHistory } from 'react-router-dom';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import './Dash.css'
import {
  PieChart,
  DataVizPalette,
  IDataPoint as FluentIDataPoint
} from '@fluentui/react-charting';
import MDR from './MDR';

// Extend IDataPoint to include 'color'
interface IDataPoint extends FluentIDataPoint {
  color: string;
}

interface IDashboardProps {
  context: WebPartContext;
}

interface IDocumentItem {
  Id: number;
  FileLeafRef: string;
  FileRef: string;
}

// Define interfaces for SharePoint list items
interface IProjectItem {
  ID: number;
  ProjectName: string;
  ProjectOverview: string;
  PreparedBy: {
    Title: string;
  };
  ProjectType: {
    ProjectType: string;
  };
  ClientName: string;
  Date: string;
  Status: string;
}

interface IApprovalItem {
  ID: number;
  ProjectCreationListIDId: number;
  DeliverablesDetailsIdId: number;
  Status: string;
  Level?: string; // Added Level property
  AssignedTo?: {
    Title: string;
  };
  ApprovalDate?: string;
  Comments?: string;
  Remarks?: string;
  ApproverRole?: string; // Added ApproverRole property
}

// Add to your existing interfaces
interface IApprovalHierarchyItem {
  ID: number;
  DeliverablesDetailsIdId: number;
  AssignedTo: {
    Title: string;
  }[];
  ApproverRole: string;
  Level: number;
  Status: string;
}

interface IDeliverableItem {
  ID: number;
  ProjectCreationListIDId: number;
  Status: string;
  DeliverablesDocumentIDId: number | null;
  Deliverables: string;
  Area: string;
  AssignedTo: {
    Title: string;
  };
  DocNumber: string;
  Organization: string;
  Modified: string;
  RevisionNumber: string;
  FileLeafRef: string;
  DocumentComments: string;
}

interface ICommentItem {
  ID: number;
  ProjectID: number;
}

interface ICard {
  id: number;
  title: string;
  type: string;
  user: string;
  statusCounts: number[];
  description: string;
  documents: number;
  comments: number;
  progress: { current: number; total: number };
}

interface IProjectDetails {
  ProjectName: string;
  ProjectOverview: string;
  PreparedBy: string;
  ProjectType: string;
  ClientName: string;
  Date: string;
  Status: string;
  ProjCreationId: number;
  CompletedTask: number;
  TotalNoComment: number;
  Members: number;
  DeliverablesDetailsArr: IDeliverableItem[];
}

const Dashboard = ({ props }: any) => {
   const elementRef = React.useRef<HTMLDivElement>(null);
    const sp: SPFI = getSP();
     const siteUrl = props.siteUrl
       const { useHide }: any = React.useContext(UserContext);
  const history = useHistory();
  const [cards, setCards] = React.useState<ICard[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [showDetails, setShowDetails] = React.useState(false);
  const [selectedProject, setSelectedProject] = React.useState<IProjectDetails | null>(null);
  const [deliverableItems, setDeliverableItems] = React.useState<IDeliverableItem[]>([]);
  const [commentItems, setCommentItems] = React.useState<ICommentItem[]>([]);
  const [documentCache, setDocumentCache] = React.useState<Map<number, IDocumentItem>>(new Map()); // Cache for documents

  const [approvalItems, setApprovalItems] = React.useState<IApprovalItem[]>([]); // Add this state
  const [showAuditHistory, setShowAuditHistory] = React.useState(false);
  const [auditHistoryData, setAuditHistoryData] = React.useState<IApprovalItem[]>([]);

  const [showApprovalHierarchy, setShowApprovalHierarchy] = React.useState(false);
  const [selectedHierarchyDeliverableId, setSelectedHierarchyDeliverableId] = React.useState<number | null>(null);
  const [approvalHierarchyData, setApprovalHierarchyData] = React.useState<IApprovalHierarchyItem[]>([]);

  // Add state for PieChart data
  const [statusChartData, setStatusChartData] = React.useState<IDataPoint[]>([]);

  // Add state for MDR view
  const [showMDR, setShowMDR] = React.useState(false);

  // Add state for NewRequest view
  const [showNewRequest, setShowNewRequest] = React.useState(false);

  React.useEffect(() => {
    const initializeSP = async () => {
      try {
        const sp = spfi().using(SPFx(props.context));
        await loadDashboardData(sp);
      } catch (error) {
        console.error("Error initializing SharePoint:", error);
        setLoading(false);
      }
    };

    initializeSP();
  }, [props.context]);

  // Update chart data when project details change
  React.useEffect(() => {
    if (selectedProject && selectedProject.DeliverablesDetailsArr.length > 0) {
      updateChartData(selectedProject.DeliverablesDetailsArr);
    }
  }, [selectedProject]);

  const updateChartData = (deliverables: IDeliverableItem[]) => {
    const statusCount: { [key: string]: number } = {};
    
    // Count deliverables by status
    deliverables.forEach(item => {
      const status = item.Status || 'Unknown';
      statusCount[status] = (statusCount[status] || 0) + 1;
    });

    // Define color mapping for different statuses
    const statusColors: { [key: string]: string } = {
      'In-Progress': DataVizPalette.color8, // Blue
      'Pending': DataVizPalette.color9,     // Orange
      'Rework': DataVizPalette.color10,     // Yellow
      'Rejected': DataVizPalette.color11,   // Red
      'Approved': DataVizPalette.color12,   // Green
      'Completed': DataVizPalette.color12,  // Green (same as Approved)
      'Unknown': DataVizPalette.color7      // Gray
    };

    // Convert to IDataPoint format for PieChart
    const chartData: IDataPoint[] = Object.entries(statusCount).map(([status, count]) => ({
      x: status,
      y: count,
      color: statusColors[status] || DataVizPalette.color7,
      legend: status
    }));

    setStatusChartData(chartData);
  };

  const loadDashboardData = async (sp: any) => {
    try {
      console.log("Loading dashboard data...");

      // Get project items
      const projectItems: IProjectItem[] = await sp.web.lists.getByTitle("ProjectCreationList").items.select(
        "*",
        "ID",
        "ProjectName",
        "ProjectOverview",
        "PreparedBy/Title",
        "ProjectType/ProjectType",
        "ClientName",
        "Date",
        "Status"
      ).expand(
        "PreparedBy",
        "ProjectType"
      )();

      console.log("Project items loaded:", projectItems);

      // Get approval items
      const approvalItems: IApprovalItem[] = await sp.web.lists.getByTitle("ProjectApprovals").items.select(
        "*",
        "AssignedTo/Title"
      ).expand(
        "AssignedTo"
      ).getAll();

      console.log("Approval items loaded:", approvalItems);
      setApprovalItems(approvalItems);

      // Get deliverable items
      const deliverables: IDeliverableItem[] = await sp.web.lists.getByTitle("DeliverablesDetails").items.select(
        "*",
        "Deliverables",
        "Area",
        "AssignedTo/Title",
        "DocNumber",
        "Organization",
        "Modified",
        "RevisionNumber",
        "FileLeafRef",
        "DocumentComments",
        "ProjectCreationListIDId",
        "Status",
        "DeliverablesDocumentIDId"
      ).expand("AssignedTo").getAll();

      console.log("Deliverable items loaded:", deliverables);
      setDeliverableItems(deliverables);

      // Get comment items
      const comments: ICommentItem[] = await sp.web.lists.getByTitle("DocumentComments").items.select(
        "*",
      )();

      console.log("Comment items loaded:", comments);
      setCommentItems(comments);

      // Transform data to cards - KEEPING THE ORIGINAL LOGIC
      const projectCards: ICard[] = projectItems.map((proj: IProjectItem) => {
        const projectDeliverables = deliverables.filter((d: IDeliverableItem) => d.ProjectCreationListIDId === proj.ID);
        const projectComments = comments.filter((c: ICommentItem) => c.ProjectID === proj.ID);

        const statusMap = ["Pending", "In-Progress", "Approved", "Rejected", "Rework"];
        const statusCounts = statusMap.map(status => {
          if (status === "Rework") {
            // Count items that have DeliverablesDocumentID and status is Pending
            return deliverables.filter((d: IDeliverableItem) =>
              d.DeliverablesDocumentIDId !== null &&
              d.DeliverablesDocumentIDId !== undefined &&
              d.Status === "Pending"
            ).length;
          } else if (status === "Pending") {
            // Count pending items that DON'T have DeliverablesDocumentIDId (not rework)
            return deliverables.filter((d: IDeliverableItem) =>
              (d.DeliverablesDocumentIDId === null ||
                d.DeliverablesDocumentIDId === undefined) &&
              d.Status === "Pending"
            ).length;
          } else {
            // For other statuses, count normally
            return deliverables.filter((d: IDeliverableItem) => d.Status === status).length;
          }
        });

        const totaldeliverables = projectDeliverables.length;
        const completed = projectDeliverables.filter((d: IDeliverableItem) => d.Status === "Approved").length;

        return {
          id: proj.ID,
          title: proj.ProjectName || "Untitled Project",
          type: proj.ProjectType?.ProjectType || "N/A",
          user: proj.PreparedBy?.Title || "Unknown User",
          statusCounts,
          description: proj.ProjectOverview || "No description available",
          documents: projectDeliverables.length,
          comments: projectComments.length,
          progress: {
            current: completed,
            total: totaldeliverables > 0 ? totaldeliverables : 1
          }
        };
      });

      setCards(projectCards);
    } catch (err) {
      console.error("Error loading dashboard data:", err);
      // Fallback with sample data for testing
      setCards([
        {
          id: 1,
          title: "Sample Project",
          type: "Development",
          user: "Test User",
          statusCounts: [1, 2, 3, 0, 1],
          description: "This is sample data",
          documents: 5,
          comments: 3,
          progress: { current: 2, total: 5 }
        }
      ]);
    } finally {
      setLoading(false);
    }
  };

  // Add this function to handle audit history
  const handleOpenAuditHistory = async (deliverableId: number) => {
    try {
      setLoading(true);

      // Filter approval items based on DeliverablesDetailsId
      const filteredApprovals = approvalItems.filter(
        approval => approval.DeliverablesDetailsIdId === deliverableId
      );

      console.log(`Audit history for deliverable ${deliverableId}:`, filteredApprovals);
      setAuditHistoryData(filteredApprovals);
      setShowAuditHistory(true);

    } catch (error) {
      console.error("Error loading audit history:", error);
      alert("Error loading audit history. Please try again.");
    } finally {
      setLoading(false);
    }
  };

  const handleCloseAuditHistory = () => {
    setShowAuditHistory(false);
    setAuditHistoryData([]);
  };

  const handleOpenHierarchy = async (deliverableId: number) => {
    try {
      setLoading(true);
      setSelectedHierarchyDeliverableId(deliverableId);

      const sp = spfi().using(SPFx(props.context));

      // Fetch approval hierarchy data based on deliverable ID
      const hierarchyData: IApprovalHierarchyItem[] = await sp.web.lists.getByTitle("ApprovalHierarchy")
        .items.filter(`DeliverablesDetailsIdId eq ${deliverableId}`)
        .select(
          "*",
          "ID",
          "DeliverablesDetailsIdId",
          "AssignedTo/Title"
        )
        .expand("AssignedTo")
        .orderBy("SerialNumber", true)();

      console.log(`Approval hierarchy for deliverable ${deliverableId}:`, hierarchyData);
      setApprovalHierarchyData(hierarchyData);
      setShowApprovalHierarchy(true);

    } catch (error) {
      console.error("Error loading approval hierarchy:", error);
      alert("Error loading approval hierarchy. Please try again.");
    } finally {
      setLoading(false);
    }
  };

  const handleCloseApprovalHierarchy = () => {
    setShowApprovalHierarchy(false);
    setSelectedHierarchyDeliverableId(null);
    setApprovalHierarchyData([]);
  };

  const getDocumentInfo = async (documentId: number | null): Promise<{ fileName: string; fileUrl: string }> => {
    if (!documentId) return { fileName: "", fileUrl: "" };

    // Check if document is already in cache
    if (documentCache.has(documentId)) {
      const doc = documentCache.get(documentId)!;
      return { fileName: doc.FileLeafRef, fileUrl: doc.FileRef };
    }

    try {
      const sp = spfi().using(SPFx(props.context));

      // Fetch only the specific document
      const document: IDocumentItem = await sp.web.lists.getByTitle("DeliverablesDocument").items.getById(documentId).select(
        "Id",
        "FileLeafRef",
        "FileRef"
      )();

      // Add to cache
      setDocumentCache(prev => new Map(prev.set(documentId, document)));

      return {
        fileName: document.FileLeafRef,
        fileUrl: document.FileRef
      };
    } catch (error) {
      console.error("Error fetching document info:", error);
      return { fileName: "", fileUrl: "" };
    }
  };

  const handleDownload = async (documentId: number | null) => {
    if (!documentId) return;

    try {
      const sp = spfi().using(SPFx(props.context));

      // Get document info first (this will use cache or fetch if not available)
      const documentInfo = await getDocumentInfo(documentId);

      if (!documentInfo.fileName) {
        throw new Error("Document not found");
      }

      // Get the file content as blob
      const fileContent = await sp.web.getFileByServerRelativePath(documentInfo.fileUrl).getBlob();

      // Create download link
      const url = window.URL.createObjectURL(fileContent);
      const a = document.createElement('a');
      a.style.display = 'none';
      a.href = url;
      a.download = documentInfo.fileName;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

    } catch (error) {
      console.error("Error downloading file:", error);
      alert("Error downloading file. Please try again.");
    }
  };

  // Create a component for the document cell to handle async loading
  const DocumentCell: React.FC<{ documentId: number | null }> = ({ documentId }) => {
    const [documentInfo, setDocumentInfo] = React.useState<{ fileName: string; fileUrl: string } | null>(null);
    const [loading, setLoading] = React.useState(false);

    React.useEffect(() => {
      const loadDocumentInfo = async () => {
        if (!documentId) return;

        setLoading(true);
        try {
          const info = await getDocumentInfo(documentId);
          setDocumentInfo(info);
        } catch (error) {
          console.error("Error loading document info:", error);
        } finally {
          setLoading(false);
        }
      };

      loadDocumentInfo();
    }, [documentId]);

    if (!documentId) {
      return null;
    }

    if (loading) {
      return <span className={styles.loadingText}>Loading...</span>;
    }

    if (!documentInfo || !documentInfo.fileName) {
      return <span className={styles.noDocument}>No document</span>;
    }

    return (
      <a
        href="#"
        onClick={(e) => {
          e.preventDefault();
          handleDownload(documentId);
        }}
        className={styles.fileLink}
        title={`Download ${documentInfo.fileName}`}
      >
        {documentInfo.fileName}
      </a>
    );
  };

  const handleCardClick = async (projectId: number) => {
    try {
      setLoading(true);
      const sp = spfi().using(SPFx(props.context));

      // Get project details
      const projectItem: IProjectItem = await sp.web.lists.getByTitle("ProjectCreationList").items.getById(projectId)
        .select(
          "ProjectName",
          "ProjectOverview",
          "PreparedBy/Title",
          "ProjectType/ProjectType",
          "ClientName",
          "Date",
          "Status"
        )
        .expand("PreparedBy", "ProjectType")();

      // Get deliverables for this project
      const projectDeliverables = deliverableItems.filter(d => d.ProjectCreationListIDId === projectId);

      // Get comments for this project
      const projectComments = commentItems.filter(c => c.ProjectID === projectId);

      // Get unique assigned team members from deliverables
      const uniqueAssignedMembers = new Set<string>();
      projectDeliverables.forEach(deliverable => {
        if (deliverable.AssignedTo?.Title) {
          uniqueAssignedMembers.add(deliverable.AssignedTo.Title);
        }
      });

      const projectDetails: IProjectDetails = {
        ProjCreationId: projectId || 0,
        ProjectName: projectItem.ProjectName || "Untitled Project",
        ProjectOverview: projectItem.ProjectOverview || "No overview available",
        PreparedBy: projectItem.PreparedBy?.Title || "Unknown",
        ProjectType: projectItem.ProjectType?.ProjectType || "N/A",
        ClientName: projectItem.ClientName || "N/A",
        Date: projectItem.Date ? new Date(projectItem.Date).toLocaleDateString() : "N/A",
        Status: projectItem.Status || "Unknown",
        CompletedTask: projectDeliverables.filter(d => d.Status === "Approved" || d.Status === "Completed").length,
        TotalNoComment: projectComments.length,
        Members: uniqueAssignedMembers.size,
        DeliverablesDetailsArr: projectDeliverables
      };

      setSelectedProject(projectDetails);
      setShowDetails(true);
    } catch (error) {
      console.error("Error loading project details:", error);
      // Fallback sample data
      setSelectedProject({
        ProjCreationId: 0,
        ProjectName: "Sample Project",
        ProjectOverview: "This is sample project overview",
        PreparedBy: "Test User",
        ProjectType: "Development",
        ClientName: "Sample Client",
        Date: new Date().toLocaleDateString(),
        Status: "In Progress",
        CompletedTask: 5,
        TotalNoComment: 3,
        Members: 4,
        DeliverablesDetailsArr: []
      });
      setShowDetails(true);
    } finally {
      setLoading(false);
    }
  };

  const handleBackToDashboard = () => {
    setShowDetails(false);
    setSelectedProject(null);
  };

  const handleViewMDR = () => {
    if (selectedProject) {
      // Get the project ID from the selected project
      const projectId = selectedProject.ProjCreationId || cards.find(card => card.title === selectedProject.ProjectName)?.id;

      if (projectId) {
        // Set the project ID in URL params for MDR component
        const url = new URL(window.location.href);
        url.searchParams.set('CreationId', projectId.toString());
        window.history.pushState({}, '', url.toString());

        // Show MDR view
        setShowMDR(true);
      } else {
        console.error("Project ID not found");
        alert("Unable to redirect: Project ID not found");
      }
    }
  };

  const getStatusBadgeClass = (status: string) => {
    switch (status) {
      case 'In-Progress': return styles.badgeInProgress;
      case 'Pending': return styles.badgePending;
      case 'Rework': return styles.badgeRework;
      case 'Rejected': return styles.badgeRejected;
      case 'Approved': return styles.badgeApproved;
      default: return styles.badgeDefault;
    }
  };

  const menuProps: IContextualMenuProps = {
    items: [
      {
        key: 'newProject',
        text: 'New Project',
        onClick: () => setShowNewRequest(true)
      }
    ]
  };

  if (loading) return <div className={styles.loading}>Loading...</div>;

  // Show dashboard view
  return (
      <div id="wrapper" ref={elementRef}>
          <div className="app-menu" id="myHeader">
            <VerticalSideBar _context={sp} />
          </div>
    
          <div className="content-page">
            <HorizontalNavbar _context={sp} siteUrl={siteUrl} />
      <div
          className="content"
          style={{
            marginLeft: `${!useHide ? "240px" : "80px"}`,

            marginTop: "0rem",
          }}
        >
          <div className='container-fluid  paddb'>

         
    <div className={styles.dashboard}>
      {showMDR ? (
        <MDR context={props.context} />
      ) : showNewRequest ? (
        <NewRequest context={props.context} onCancel={() => setShowNewRequest(false)} />
      ) : showDetails && selectedProject ? (
        <div className={styles.projectDetails}>
          {/* Horizontal Scroll Wrapper for Entire Details View */}
          <div className={styles.detailsScrollWrapper}>

            {/* Header */}
            <div className={styles.pageTitleRow}>
              <div className={styles.pageTitleBox}>
                <h4 className={styles.pageTitle}>Folder__{selectedProject.ProjectName}</h4>
                <div className={styles.pageTitleRight}>
                  <button type="button" onClick={handleViewMDR} className={styles.btnDark}>
                    View MDR
                  </button>
                  <button type="button" onClick={handleBackToDashboard} className='btnDark1'>
                    Back
                  </button>
                </div>
              </div>
            </div>

            {/* Stats Cards */}
            <div className={styles.statsRow}>
              <div className={styles.statsCol}>
                <div className={styles.statsCard}>
                  <div className={styles.cardBody}>
                    <div className={styles.statsRowInner}>
                      <div className={styles.statsIconCol}>
                        <div className={`${styles.avatarMd} ${styles.bgSuccess}`}>
                          <i className={`${styles.feAward} ${styles.avatarTitle}`}>✓</i>
                        </div>
                      </div>
                      <div className={styles.statsTextCol}>
                        <div className={styles.textEnd}>
                          <h3 className={styles.statsNumber}>{selectedProject.CompletedTask}</h3>
                          <p className={styles.statsLabel}>Document(s)</p>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              <div className={styles.statsCol}>
                <div className={styles.statsCard}>
                  <div className={styles.cardBody}>
                    <div className={styles.statsRowInner}>
                      <div className={styles.statsIconCol}>
                        <div className={`${styles.avatarMd} ${styles.bgDanger}`}>
                          <i className={`${styles.feMessageSquare} ${styles.avatarTitle}`}>💬</i>
                        </div>
                      </div>
                      <div className={styles.statsTextCol}>
                        <div className={styles.textEnd}>
                          <h3 className={styles.statsNumber}>{selectedProject.TotalNoComment}</h3>
                          <p className={styles.statsLabel}>Comment(s)</p>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              <div className={styles.statsCol}>
                <div className={styles.statsCard}>
                  <div className={styles.cardBody}>
                    <div className={styles.statsRowInner}>
                      <div className={styles.statsIconCol}>
                        <div className={`${styles.avatarMd} ${styles.bgWarning}`}>
                          <i className={`${styles.feUsers} ${styles.avatarTitle}`}>👥</i>
                        </div>
                      </div>
                      <div className={styles.statsTextCol}>
                        <div className={styles.textEnd}>
                          <h3 className={styles.statsNumber}>{selectedProject.Members}</h3>
                          <p className={styles.statsLabel}>Member(s)</p>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>

            {/* Project Details */}
            <div className={styles.detailsRow}>
              <div className={styles.projectInfoCol}>
                <div className={styles.projectCard}>
                  <div className={styles.cardBody}>
                    <h5 className={styles.sectionTitle}>Overview:</h5>
                    <p className={styles.projectOverview}>{selectedProject.ProjectOverview}</p>

                    <div className={styles.detailsGrid}>
                      <div className={styles.detailItem}>
                        <h5 className={styles.detailTitle}>Project Date</h5>
                        <p>{selectedProject.Date}</p>
                      </div>

                      <div className={styles.detailItem}>
                        <h5 className={styles.detailTitle}>Prepared By</h5>
                        <p>{selectedProject.PreparedBy}</p>
                      </div>

                      <div className={styles.detailItem}>
                        <h5 className={styles.detailTitle}>Project Type</h5>
                        <p>{selectedProject.ProjectType}</p>
                      </div>

                      <div className={styles.detailItem}>
                        <h5 className={styles.detailTitle}>Client Name</h5>
                        <p>{selectedProject.ClientName}</p>
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              <div className={styles.chartCol}>
                <div className={styles.chartCard}>
                  <div className={styles.cardBody}>
                    <h4 style={{fontWeight:'bold'}} className='header-title line18 font-8 text-dark newtextdark fw-bold mb-0'>Deliverables by Status</h4>
                    <div className={styles.chartContainer}>
                      {statusChartData.length > 0 ? (
                        <div className={styles.pieChartWrapper}>
                          <PieChart
                            data={statusChartData}
                            width={250}
                            height={250}
                           // chartTitle="Deliverables by Status"
                            // hideLabels={false}
                            // showLabelsInPercent={true}
                          />
                          <div className={styles.chartLegend}>
                            {statusChartData.map((item, index) => (
                              <div key={index} className={styles.legendItem}>
                                <div
                                  className={styles.legendColor}
                                  style={{ backgroundColor: item.color }}
                                ></div>
                                <span className={styles.legendText}>
                                  {item.x}: {item.y}
                                </span>
                              </div>
                            ))}
                          </div>
                        </div>
                      ) : (
                        <div className={styles.chartPlaceholder}>
                          <div className={styles.chartStats}>
                            <div className={styles.chartStat}>
                              <span className={styles.statCount}>
                                {selectedProject.DeliverablesDetailsArr.filter(d => d.Status === 'Approved' || d.Status === 'Completed').length}
                              </span>
                              <span className={styles.statLabel}>Completed</span>
                            </div>
                            <div className={styles.chartStat}>
                              <span className={styles.statCount}>
                                {selectedProject.DeliverablesDetailsArr.filter(d => d.Status === 'In-Progress').length}
                              </span>
                              <span className={styles.statLabel}>In Progress</span>
                            </div>
                            <div className={styles.chartStat}>
                              <span className={styles.statCount}>
                                {selectedProject.DeliverablesDetailsArr.filter(d => d.Status === 'Pending').length}
                              </span>
                              <span className={styles.statLabel}>Pending</span>
                            </div>
                          </div>
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              </div>
            </div>

            {/* Deliverables Table */}
            <div className={styles.deliverablesSection}>
              <div className={styles.deliverablesCard}>
                <div className={styles.cardBody}>
                  <h4 className='header-title line18 font-8 text-dark newtextdark fw-bold mb-0'>Deliverables Status</h4>

                  <div className={styles.tableContainer}>
                    <table className={styles.deliverablesTable}>
                      <thead>
                        <tr>
                          <th>Deliverables</th>
                          <th>Area</th>
                          <th style={{minWidth:'100px'}}>Assigned To</th>
                          <th>Document Number</th>
                          <th>Organization</th>
                          <th>Action Date Time</th>
                          <th style={{minWidth:'100px'}}>Revision Number</th>
                          <th style={{minWidth:'100px'}}>Attachment</th>
                          <th style={{minWidth:'100px'}}>Status</th>
                          <th>Remark</th>
                          <th style={{minWidth:'100px'}}>Audit History</th>
                          <th style={{minWidth:'100px'}}>Approval Hierarchy</th>
                        </tr>
                      </thead>
                      <tbody>
                        {selectedProject.DeliverablesDetailsArr.map((deliverable) => (
                          <tr key={deliverable.ID}>
                            <td>{deliverable.Deliverables}</td>
                            <td>{deliverable.Area}</td>
                            <td style={{minWidth:'100px'}}>{deliverable.AssignedTo?.Title}</td>
                            <td>{deliverable.DocNumber}</td>
                            <td>{deliverable.Organization}</td>
                            <td>
                              {deliverable.Status !== 'Pending' && deliverable.Modified
                                ? new Date(deliverable.Modified).toLocaleString()
                                : ''}
                            </td>
                            <td style={{minWidth:'100px'}} className={styles.textCenter}>{deliverable.RevisionNumber}</td>
                            <td style={{minWidth:'100px'}}>
                              <DocumentCell documentId={deliverable.DeliverablesDocumentIDId} />
                            </td>
                            <td style={{minWidth:'100px'}}>
                              <span className={`${styles.statusBadge} ${getStatusBadgeClass(deliverable.Status)}`}>
                                {deliverable.Status}
                              </span>
                            </td>
                            <td>{deliverable.DocumentComments}</td>
                            <td style={{minWidth:'100px'}} className={styles.textCenter}>
                              <button style={{minWidth:'40px'}}
                                type="button"
                                className={styles.btnOutlineSuccess}
                                onClick={() => handleOpenAuditHistory(deliverable.ID)}
                              >
                                🔍
                              </button>
                            </td>
                            <td style={{minWidth:'100px'}} className={styles.textCenter}>
                              <button style={{minWidth:'40px'}}
                                type="button"
                                className={styles.btnOutlineDanger}
                                onClick={() => handleOpenHierarchy(deliverable.ID)}
                              >
                                👁️
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    {selectedProject.DeliverablesDetailsArr.length === 0 && (
                      <div className={styles.noData}>
                        <p>No deliverables found for this project.</p>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </div>

          </div> {/* End of detailsScrollWrapper */}

          {/* AUDIT HISTORY MODAL */}
          {showAuditHistory && (
            <div className={styles.modalOverlay}>
              <div className={styles.modalContent}>
                <div className={styles.modalHeader}>
                  <h3>Audit History</h3>
                  <button
                    type="button"
                    className={styles.closeButton}
                    onClick={handleCloseAuditHistory}
                  >
                    ×
                  </button>
                </div>

                <div className={styles.modalBody}>
                  {auditHistoryData.length > 0 ? (
                    <div className={styles.tableContainer}>
                      <table className={styles.auditTable}>
                        <thead>
                          <tr>
                            <th>Approval Level</th>
                            <th>Approval Role</th>
                            <th>Approver</th>
                            <th>Approval Date</th>
                            <th>Status</th>
                            <th>Remarks</th>
                          </tr>
                        </thead>
                        <tbody>
                          {auditHistoryData.map((approval) => (
                            <tr key={approval.ID}>
                              <td>{approval.Level}</td>
                              <td>{approval.ApproverRole}</td>
                              <td>{approval.AssignedTo?.Title || 'N/A'}</td>
                              <td>
                                {approval.ApprovalDate
                                  ? new Date(approval.ApprovalDate).toLocaleDateString()
                                  : 'N/A'
                                }
                              </td>
                              <td>
                                <span className={`${styles.statusBadge} ${getStatusBadgeClass(approval.Status)}`}>
                                  {approval.Status}
                                </span>
                              </td>
                              <td>{approval.Remarks || 'No comments'}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  ) : (
                    <div className={styles.noData}>
                      <p>No audit history found for this deliverable.</p>
                    </div>
                  )}
                </div>

                <div className={styles.modalFooter}>
                  <button
                    type="button"
                    className={styles.btnDark}
                    onClick={handleCloseAuditHistory}
                  >
                    Close
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* Approval Hierarchy Modal */}
          {showApprovalHierarchy && (
            <div className={styles.modalOverlay}>
              <div className={styles.modalContent}>
                <div className={styles.modalHeader}>
                  <h3>Approval Hierarchy</h3>
                  <button
                    type="button"
                    className={styles.closeButton}
                    onClick={handleCloseApprovalHierarchy}
                  >
                    ×
                  </button>
                </div>

                <div className={styles.modalBody}>
                  <div className={styles.auditInfo}>
                    <h4>Deliverable ID: {selectedHierarchyDeliverableId}</h4>
                    <p>Total Approval Steps: {approvalHierarchyData.length}</p>
                  </div>

                  {approvalHierarchyData.length > 0 ? (
                    <div className={styles.tableContainer}>
                      <table className={styles.auditTable}>
                        <thead>
                          <tr>
                            <th>Approval Level</th>
                            <th>Approval Role</th>
                            <th>Approver</th>
                            <th>Status</th>
                          </tr>
                        </thead>
                        <tbody>
                          {approvalHierarchyData.map((hierarchy) => (
                            <tr key={hierarchy.ID}>
                              <td className={styles.textCenter}>{hierarchy.Level}</td>
                              <td>{hierarchy.ApproverRole}</td>
                              <td>
                                {hierarchy.AssignedTo && hierarchy.AssignedTo.length > 0 ? (
                                  <div className={styles.userList}>
                                    {hierarchy.AssignedTo.map((user, index) => (
                                      <div key={index} className={styles.userItem}>
                                        {user.Title}
                                        {index < hierarchy.AssignedTo.length - 1 && ', '}
                                      </div>
                                    ))}
                                  </div>
                                ) : (
                                  'N/A'
                                )}
                              </td>
                              <td>
                                <span className={`${styles.statusBadge} ${getStatusBadgeClass(hierarchy.Status)}`}>
                                  {hierarchy.Status}
                                </span>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  ) : (
                    <div className={styles.noData}>
                      <p>No approval hierarchy found for this deliverable.</p>
                    </div>
                  )}
                </div>

                <div className={styles.modalFooter}>
                  <button
                    type="button"
                    className={styles.btnDark}
                    onClick={handleCloseApprovalHierarchy}
                  >
                    Close
                  </button>
                </div>
              </div>
            </div>
          )}
        </div>
      ) : (
        <>
          <div className={styles.controls}>
            <h2 style={{margin:'0px',fontWeight:'600'}} className='fw-bold text-dark header-title'>Dashboard</h2>

            <div className={styles.searchArea}>
              <input
                type="text"
                placeholder="Search Card..."
                className={styles.searchInput}
              />
              <button className={styles.filterButton}>All</button>
              <DefaultButton
                text="New Request"
                menuProps={menuProps}
                className={styles.newRequest}
              />
            </div>
          </div>

          <div className={styles.cardsGrid}>
            {cards.length > 0 ? (
              cards.map((card, index) => (
                <div
                  key={index}
                  className={styles.card}
                  onClick={() => handleCardClick(card.id)}
                  style={{ cursor: 'pointer' }}
                >
                  <div className={styles.cardHeader}>
                    <h3>{card.title}</h3>
                    <span>...</span>
                  </div>
                  <span className={styles.type}>{card.type}</span>
                  <div className={styles.userInfo}>
                    <span>👤 {card.user}</span>
                  </div>
                  <div className={styles.statusCounts}>
                    {card.statusCounts.map((count, idx) => (
                      <span key={idx} className={styles.statusBadge}>{count}</span>
                    ))}
                  </div>
                  <p className={styles.description}>{card.description}</p>
                  <div className={styles.stats}>
                    <span>📄 {card.documents} Documents</span>
                    <span>💬 {card.comments} Comments</span>
                  </div>
                  <div className={styles.progress}>
                    <span>Documents</span>
                    <span>{card.progress.current}/{card.progress.total}</span>
                    <div className={styles.progressBar}>
                      <div
                        className={styles.progressFill}
                        style={{ width: `${(card.progress.current / card.progress.total) * 100}%` }}
                      ></div>
                    </div>
                  </div>
                </div>
              ))
            ) : (
              <div className={styles.noData}>
                <p>No projects found. Please check your SharePoint list configuration.</p>
              </div>
            )}
          </div>
        </>
      )}
    </div>
    </div>
          </div>
      </div> </div>
  );
};

const MyProjects: React.FC<IMyProjectProps> = (props) => (
  <Provider>
    <Dashboard props={props} />
  </Provider>
);

export default MyProjects;