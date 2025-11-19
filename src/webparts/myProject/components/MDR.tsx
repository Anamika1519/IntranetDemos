import * as React from 'react';
// import { useLocation, useHistory } from 'react-router-dom';
import { useHistory } from 'react-router-dom';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
// import { ExcelExportService } from '../services/ExcelExportService';
import styles from './MDR.module.scss';
 
interface IMDRProps {
    context: WebPartContext;
}
 
interface IProjectItem {
    ID: number;
    ProjectName: string;
    ProjectType: {
        ProjectType: string;
    };
}
 
interface IDeliverableItem {
    Id: number;
    DocNumber: string;
    Deliverables: string;
    Area: string;
    Organization: string;
    RevisionNumber: number;
    DocumentComments: string;
    DocumentType: string;
    Status: string;
    ProjectCreationListID: {
        ProjectName: string;
    };
    AssignedTo: {
        Title: string;
    };
    IncomingRev: string[];
    IncomingStatus: string[];
    IncomingSubmitDate: string[];
    IncomingTransmittalNo: string[];
    OutgoingReturnDate: string[];
    OutgoingTransmittalNo: string[];
    OutgoingCodeStatus: string[];
    ServerLink: string[];
    FileName: string[];
}
 
interface IApprovalItem {
    ID: number;
    RevisionNumber: number;
    RequestedRole: string;
    IncomingDate: string;
    OutgoingDate: string;
    ApproverRole: string;
}
 
interface IDocumentItem {
    FileRef: string;
    FileLeafRef: string;
}
 
const MDR: React.FC<IMDRProps> = ({ context }) => {
    //   const location = useLocation();
    const history = useHistory();
    //   const [creationId, setCreationId] = React.useState<number | null>(null);
    const [projectName, setProjectName] = React.useState<string>('');
    const [deliverablesDetails, setDeliverablesDetails] = React.useState<IDeliverableItem[]>([]);
    const [submissions, setSubmissions] = React.useState<number[]>([]);
    //   const [revisionNumber, setRevisionNumber] = React.useState<number>(0);
    const [loading, setLoading] = React.useState<boolean>(true);
     
    // this was old working
    // React.useEffect(() => {
    //     const getCreationIdFromUrl = () => {
    //         try {
    //             const hash = window.location.hash;
    //             if (hash) {
    //                 const params = new URLSearchParams(hash.split('#')[2] || hash.split('?')[1]);
    //                 const creationIdParam = params.get('CreationId');
    //                 console.log('CreationId:', creationIdParam);
    //                 return creationIdParam ? parseInt(creationIdParam, 10) : null;
    //             }
    //         } catch (error) {
    //             console.error('Error parsing URL:', error);
    //         }
    //         return null;
    //     };
 
    //     const creationIdFromUrl = getCreationIdFromUrl();
    //     if (creationIdFromUrl) {
    //         //   setCreationId(creationIdFromUrl);
    //         getCreationDetails(creationIdFromUrl);
    //     } else {
    //         setLoading(false);
    //     }
    // }, []);
  React.useEffect(() => {
        const getCreationIdFromUrl = () => {
            try {
                let params: URLSearchParams;
 
                if (window.location.search) {
                    params = new URLSearchParams(window.location.search);
                }
                // Fallback for hash-style URLs (#/?CreationId=)
                else if (window.location.hash) {
                    const hash = window.location.hash;
                    params = new URLSearchParams(hash.split('?')[1] || '');
                } else {
                    return null;
                }
 
                const creationIdParam = params.get('CreationId');
                console.log('CreationId:', creationIdParam);
                return creationIdParam ? parseInt(creationIdParam, 10) : null;
            } catch (error) {
                console.error('Error parsing URL:', error);
                return null;
            }
        };
 
        const creationIdFromUrl = getCreationIdFromUrl();
        if (creationIdFromUrl) {
            // setCreationId(creationIdFromUrl);
            getCreationDetails(creationIdFromUrl);
        } else {
            setLoading(false);
        }
    }, []);
    const getCreationDetails = async (creationId: number) => {
        setLoading(true);
        try {
            const sp = spfi().using(SPFx(context));
 
            // Get project details
            const projectItems: IProjectItem[] = await sp.web.lists.getByTitle("ProjectCreationList")
                .items
                .select("ID", "ProjectName", "ProjectType/ProjectType")
                .expand("ProjectType")
                .filter(`ID eq ${creationId}`)
                .orderBy("Created", false)();
 
            if (projectItems.length > 0) {
                setProjectName(projectItems[0].ProjectName);
                await getDeliverablesDetails(creationId);
            } else {
                setLoading(false);
            }
        } catch (error) {
            console.error('Error retrieving project details:', error);
            setLoading(false);
        }
    };
 
    const getDocumentUrl = async (creationId: number, deliverableId: number, revision: number): Promise<{ url: string; name: string }> => {
        try {
            const sp = spfi().using(SPFx(context));
 
            const documents: IDocumentItem[] = await sp.web.lists.getByTitle("DeliverablesDocument")
                .items
                .select("FileRef", "FileLeafRef")
                .filter(`ProjectID eq ${creationId} and DeliverablesDetailsId eq ${deliverableId} and Revision eq ${revision}`)
                .orderBy("Created", false)
                .top(1)();
 
            if (documents.length > 0) {
                return {
                    url: documents[0].FileRef,
                    name: documents[0].FileLeafRef
                };
            } else {
                return { url: "", name: "" };
            }
        } catch (error) {
            console.error("Error fetching document link:", error);
            return { url: "", name: "" };
        }
    };
 
    const getDeliverablesDetails = async (creationId: number) => {
        try {
            const sp = spfi().using(SPFx(context));
 
            // Get deliverables
            const deliverables: IDeliverableItem[] = await sp.web.lists.getByTitle("DeliverablesDetails")
                .items
                .select(
                    "Id", "DocNumber", "Deliverables", "Area", "Organization",
                    "RevisionNumber", "DocumentComments", "DocumentType", "Status",
                    "ProjectCreationListID/ProjectName", "AssignedTo/Title"
                )
                .expand("ProjectCreationListID", "AssignedTo")
                .filter(`ProjectCreationListID/Id eq ${creationId}`)
                .orderBy("Created", false)();
 
            if (deliverables.length > 0) {
                // Calculate highest revision number
                const highestRevision = Math.max(...deliverables.map(item => item.RevisionNumber || 0));
                // setRevisionNumber(highestRevision);
 
                // Create submissions array
                const submissionsArray = [];
                for (let i = 0; i <= highestRevision; i++) {
                    submissionsArray.push(i);
                }
                setSubmissions(submissionsArray);
 
                // Process each deliverable
                const processedDeliverables = await Promise.all(
                    deliverables.map(async (item) => {
                        const processedItem: IDeliverableItem = {
                            ...item,
                            IncomingRev: [],
                            IncomingStatus: [],
                            IncomingSubmitDate: [],
                            IncomingTransmittalNo: [],
                            OutgoingReturnDate: [],
                            OutgoingTransmittalNo: [],
                            OutgoingCodeStatus: [],
                            ServerLink: [],
                            FileName: []
                        };
 
                        try {
                            // Get approvals for this deliverable
                            const approvals: IApprovalItem[] = await sp.web.lists.getByTitle("ProjectApprovals")
                                .items
                                .select("*", "ID", "RevisionNumber", "RequestedRole", "IncomingDate", "OutgoingDate", "ApproverRole")
                                .filter(`ProjectCreationListID/Id eq ${creationId} and DeliverablesDetailsId/Id eq ${item.Id} and ApproverRole eq 'Document Controller'`)
                                .orderBy("RevisionNumber", true)();
 
                            // Process each revision
                            for (let i = 0; i <= highestRevision; i++) {
                                if (i <= item.RevisionNumber) {
                                    const revDisplay = i === 0 ? "0" : i.toString();
                                    processedItem.IncomingRev.push(revDisplay);
 
                                    const matchingApprovals = approvals.filter(app => Number(app.RevisionNumber) === i);
 
                                    // Incoming Section
                                    const incoming = matchingApprovals.find(app => app.RequestedRole === 'Vendor' && app.IncomingDate);
                                    if (incoming) {
                                        processedItem.IncomingSubmitDate.push(incoming.IncomingDate);
                                        processedItem.IncomingTransmittalNo.push(`TR-00${i + 1}`);
                                        processedItem.IncomingStatus.push(item.Status || '-');
                                    } else {
                                        processedItem.IncomingSubmitDate.push('-');
                                        processedItem.IncomingTransmittalNo.push('-');
                                        processedItem.IncomingStatus.push('-');
                                    }
 
                                    // Outgoing Section
                                    const outgoing = matchingApprovals.find(app => app.OutgoingDate);
                                    if (outgoing) {
                                        processedItem.OutgoingReturnDate.push(outgoing.OutgoingDate);
                                        processedItem.OutgoingTransmittalNo.push(`TR-00${i + 1}`);
                                        processedItem.OutgoingCodeStatus.push(item.Status || '-');
                                    } else {
                                        processedItem.OutgoingReturnDate.push('-');
                                        processedItem.OutgoingTransmittalNo.push('-');
                                        processedItem.OutgoingCodeStatus.push('-');
                                    }
 
                                    // Document Link
                                    const doc = await getDocumentUrl(creationId, item.Id, i);
                                    processedItem.ServerLink.push(doc.url);
                                    processedItem.FileName.push(doc.name ? doc.name.replace(/^6_\d+_/, '') : '');
                                } else {
                                    // Fill with '-' for non-applicable revisions
                                    processedItem.IncomingRev.push('-');
                                    processedItem.IncomingStatus.push('-');
                                    processedItem.IncomingSubmitDate.push('-');
                                    processedItem.IncomingTransmittalNo.push('-');
                                    processedItem.OutgoingReturnDate.push('-');
                                    processedItem.OutgoingTransmittalNo.push('-');
                                    processedItem.OutgoingCodeStatus.push('-');
                                    processedItem.ServerLink.push('');
                                    processedItem.FileName.push('');
                                }
                            }
                        } catch (error) {
                            console.error(`Error processing approvals for deliverable ${item.Id}:`, error);
                        }
 
                        return processedItem;
                    })
                );
 
                setDeliverablesDetails(processedDeliverables);
            }
        } catch (error) {
            console.error('Error retrieving deliverables details:', error);
        } finally {
            setLoading(false);
        }
    };
 
    // const goToPage = () => {
    //     window.location.href = "https://officeindia.sharepoint.com/sites/ESSAProjectDocuments/SitePages/NewApp.aspx#/";
    // };
 
    const getOrdinal = (n: number): string => {
        if (n === 1) return 'st';
        if (n === 2) return 'nd';
        if (n === 3) return 'rd';
        return 'th';
    };
 
    const getStatusBadgeClass = (status: string): string => {
        switch (status) {
            case 'In-Progress': return styles.badgeInProgress;
            case 'Pending': return styles.badgePending;
            case 'Rework': return styles.badgeRework;
            case 'Rejected': return styles.badgeRejected;
            case 'Approved': return styles.badgeApproved;
            default: return styles.badgeDefault;
        }
    };
 
    const formatDate = (dateString: string): string => {
        if (!dateString || dateString === '-') return '-';
        try {
            const date = new Date(dateString);
            return date.toLocaleDateString('en-GB'); // dd-mm-yyyy format
        } catch (error) {
            return dateString;
        }
    };
 
    const handleBack = () => {
        history.goBack();
    };
 
    const exportMDRToExcel = () => {
        const table = document.querySelector(`.${styles.newTable}`) as HTMLElement;
        if (!table) return;
 
        // const tableHTML = table.outerHTML
        //     .replace(/ /g, '%20')
        //     .replace(/#/g, '%23'); // Encode colors
 
        const html = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office"
          xmlns:x="urn:schemas-microsoft-com:office:excel"
          xmlns="http://www.w3.org/TR/REC-html40">
      <head>
        <meta charset="UTF-8">
        <style>
          ${getComputedStyleCSS()}
        </style>
      </head>
      <body>
        ${table.outerHTML}
      </body>
    </html>`;
 
        const blob = new Blob([html], { type: 'application/vnd.ms-excel' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `MDR_Export_${new Date().toISOString().slice(0, 10)}.xls`;
        a.click();
        URL.revokeObjectURL(url);
    };
 
    // Helper to include all global CSS used in your page
    function getComputedStyleCSS() {
        let css = '';
        for (const sheet of Array.from(document.styleSheets)) {
            try {
                for (const rule of Array.from(sheet.cssRules || [])) {
                    css += rule.cssText;
                }
            } catch (e) {
                // Ignore CORS-restricted stylesheets
            }
        }
        return css;
    }
 
 
    // const exportMDRToExcel = async () => {
    //     try {
    //         //   await ExcelExportService.exportTableToExcel('.newtabletf', 'DeliverablesDetails.xlsx');
    //     } catch (error) {
    //         console.error('Error exporting to Excel:', error);
    //         // Fallback to basic export
    //         exportMDRToExcelBasic();
    //     }
    // };
 
    // const exportMDRToExcelBasic = () => {
    //     // Basic export implementation
    //     console.log('Basic export functionality');
    //     // You can implement a simple export here or use a library
    // };
 
    if (loading) {
        return (
            <div className={styles.mdrContainer}>
                <div className={styles.loading}>
                    <div className={styles.loadingSpinner}></div>
                    Loading MDR data...
                </div>
            </div>
        );
    }
 
    return (
        <div className={styles.mdrContainer}>
            <div className={styles.pageHeader}>
                <h2>MDR - {projectName || 'Project Details'}</h2>
                <div className={styles.headerActions}>
                    <button onClick={exportMDRToExcel} className={styles.btnExport}>
                        Export to Excel
                    </button>
                    {/* <button onClick={goToPage} className={styles.btnSecondary}>
                        Go to Home
                    </button> */}
                    <button onClick={handleBack} className={styles.btnDark}>
                        Back
                    </button>
                </div>
            </div>
 
            {/* MDR Table */}
            <div className={styles.mdrTableSection}>
                <div className={styles.card}>
                    <div className={styles.cardBody}>
                        <div className={styles.tableResponsive}>
                            <table className={styles.newTable}>
                                <thead>
                                    <tr>
                                        <th rowSpan={3} className={styles.tableHeader}>SNo</th>
                                        <th rowSpan={3} className={styles.tableHeader}>Project Name</th>
                                        <th rowSpan={3} className={styles.tableHeader}>Document No</th>
                                        <th rowSpan={3} className={styles.tableHeader}>Deliverables</th>
                                        <th rowSpan={3} className={styles.tableHeader}>Area</th>
                                        <th rowSpan={3} className={styles.tableHeader}>Organization</th>
                                        <th colSpan={2} rowSpan={2} className={styles.tableHeader}>Current Status</th>
                                        <th rowSpan={3} className={styles.tableHeader}>Document Type</th>
                                        {submissions.map((_, i) => (
                                            <th key={i} colSpan={8} className={styles.submissionHeader}>
                                                {i + 1}{getOrdinal(i + 1)} Submitted
                                            </th>
                                        ))}
                                    </tr>
                                    <tr>
                                        {submissions.map((_, i) => (
                                            <React.Fragment key={i}>
                                                <th colSpan={4} className={styles.incomingHeader}>Incoming</th>
                                                <th colSpan={4} className={styles.outgoingHeader}>Outgoing</th>
                                            </React.Fragment>
                                        ))}
                                    </tr>
                                    <tr>
                                        <th className={styles.latestRevision}>Latest Revision</th>
                                        <th className={styles.commentHeader}>Comment</th>
                                        {submissions.map((_, i) => (
                                            <React.Fragment key={i}>
                                                <th className={styles.revHeader}>Rev</th>
                                                <th className={styles.statusHeader}>Status</th>
                                                <th className={styles.dateHeader}>Submit Date</th>
                                                <th className={styles.transmittalHeader}>Transmittal No</th>
                                                <th className={styles.dateHeader}>Return Date</th>
                                                <th className={styles.transmittalHeader}>Transmittal No</th>
                                                <th className={styles.statusHeader}>Code Status</th>
                                                <th className={styles.documentHeader}>Document Link</th>
                                            </React.Fragment>
                                        ))}
                                    </tr>
                                </thead>
                                <tbody>
                                    {deliverablesDetails.map((item, index) => (
                                        <tr key={item.Id}>
                                            <td className={styles.tableCell}>{index + 1}</td>
                                            <td className={styles.tableCell}>{item.ProjectCreationListID?.ProjectName}</td>
                                            <td className={styles.tableCell}>{item.DocNumber}</td>
                                            <td className={styles.tableCell}>{item.Deliverables}</td>
                                            <td className={styles.tableCell}>{item.Area}</td>
                                            <td className={styles.tableCell}>{item.Organization}</td>
                                            <td className={styles.tableCell}>
                                                <div
                                                    className={styles.revisionCircle}
                                                    style={item.RevisionNumber ? {
                                                        backgroundColor: '#e2e8ff',
                                                        color: '#7989c7'
                                                    } : {}}
                                                >
                                                    {item.RevisionNumber}
                                                </div>
                                            </td>
                                            <td className={styles.tableCell}>{item.DocumentComments}</td>
                                            <td className={styles.tableCell}>{item.DocumentType}</td>
 
                                            {submissions.map((_, submissionIndex) => {
                                                const incomingRev = item.IncomingRev[submissionIndex] || '-';
                                                const incomingStatus = item.IncomingStatus[submissionIndex] || '-';
                                                const incomingSubmitDate = item.IncomingSubmitDate[submissionIndex] || '-';
                                                const incomingTransmittalNo = item.IncomingTransmittalNo[submissionIndex] || '-';
                                                const outgoingReturnDate = item.OutgoingReturnDate[submissionIndex] || '-';
                                                const outgoingTransmittalNo = item.OutgoingTransmittalNo[submissionIndex] || '-';
                                                const outgoingCodeStatus = item.OutgoingCodeStatus[submissionIndex] || '-';
                                                const serverLink = item.ServerLink[submissionIndex] || '';
                                                const fileName = item.FileName[submissionIndex] || '';
 
                                                return (
                                                    <React.Fragment key={submissionIndex}>
                                                        <td className={styles.tableCell}>
                                                            <div
                                                                className={styles.revisionCircle}
                                                                style={incomingRev !== '-' ? {
                                                                    backgroundColor: '#e2e8ff',
                                                                    color: '#7989c7'
                                                                } : {}}
                                                            >
                                                                {incomingRev}
                                                            </div>
                                                        </td>
                                                        <td className={styles.tableCell}>
                                                            {incomingStatus !== '-' ? (
                                                                <span className={`${styles.badge} ${getStatusBadgeClass(incomingStatus)}`}>
                                                                    {incomingStatus}
                                                                </span>
                                                            ) : (
                                                                '-'
                                                            )}
                                                        </td>
                                                        <td className={styles.tableCell}>
                                                            <div
                                                                className={styles.dateCircle}
                                                                style={incomingSubmitDate !== '-' ? {
                                                                    backgroundColor: '#c7eecf',
                                                                    color: '#000'
                                                                } : {}}
                                                            >
                                                                {formatDate(incomingSubmitDate)}
                                                            </div>
                                                        </td>
                                                        <td className={styles.tableCell}>{incomingTransmittalNo}</td>
                                                        <td className={styles.tableCell}>
                                                            <div className={styles.returnDateCircle}>
                                                                {formatDate(outgoingReturnDate)}
                                                            </div>
                                                        </td>
                                                        <td className={styles.tableCell}>{outgoingTransmittalNo}</td>
                                                        <td className={styles.tableCell}>
                                                            {outgoingCodeStatus !== '-' ? (
                                                                <span className={`${styles.badge} ${getStatusBadgeClass(outgoingCodeStatus)}`}>
                                                                    {outgoingCodeStatus}
                                                                </span>
                                                            ) : (
                                                                '-'
                                                            )}
                                                        </td>
                                                        <td className={styles.tableCell}>
                                                            {fileName ? (
                                                                <a
                                                                    href={serverLink}
                                                                    target="_blank"
                                                                    rel="noopener noreferrer"
                                                                    className={styles.documentLink}
                                                                    title={fileName}
                                                                >
                                                                    Link
                                                                </a>
                                                            ) : (
                                                                '-'
                                                            )}
                                                        </td>
                                                    </React.Fragment>
                                                );
                                            })}
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
 
                            {deliverablesDetails.length === 0 && (
                                <div className={styles.noData}>
                                    <p>No deliverables found for this project.</p>
                                </div>
                            )}
                        </div>
                    </div>
                </div>
            </div>
        </div>
    );
};
 
export default MDR;










