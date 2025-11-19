import * as React from 'react';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import styles from './NewRequest.module.scss';

initializeIcons();

import {
    TextField,
    Dropdown,
    DatePicker,
    IconButton,
    PrimaryButton,
    DefaultButton,
    IDropdownOption
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import Select from "react-select";
//import { useHistory } from 'react-router-dom';
import CustomPopup from './CustomPopup';
import './Newreg.css';
interface INewRequestProps {
    context: WebPartContext;
    onCancel?: () => void;
    // description: string;
    // siteUrl: string;
    // userDisplayName: string;
    // isDarkTheme: boolean;
    // environmentMessage: string;
    // hasTeamsContext: boolean;
}

interface UserOption {
    value: number;
    label: string;
    email: string;
}

// Define master data interfaces
interface AreaMaster {
    ID: number;
    Area: string;
    AreaIdentificationCode: string;
    AreaCode: string;
}

interface OrganizationMaster {
    ID: number;
    Organization: string;
    OrganizationIdentificationCode: string;
}

interface DocumentTypeMaster {
    ID: number;
    DocumentType: string;
    DocumentTypeCode: string;
}

interface DeliverablesMaster {
    ID: number;
    Deliverables: string;
}

const NewRequest: React.FC<INewRequestProps> = (props) => {
    //const history = useHistory();
    const sp: SPFI = spfi().using(SPFx(props.context));
    const [projectTypeOptions, setProjectTypeOptions] = React.useState<IDropdownOption[]>([]);
    const [AreaOptions, setAreaOptions] = React.useState<IDropdownOption[]>([]);
    const [DeliverablesOptions, setDeliverablesOptions] = React.useState<IDropdownOption[]>([]);
    const [OrganizationOptions, setOrganizationOptions] = React.useState<IDropdownOption[]>([]);
    const [DocumentTypeOptions, setDocumentTypeOptions] = React.useState<IDropdownOption[]>([]);
    const [alertMessage, setAlertMessage] = React.useState<string | null>(null);
    const [userOptions, setUserOptions] = React.useState<UserOption[]>([]);
    const [selectedUser, setSelectedUser] = React.useState<UserOption | null>(null);

    // Master data arrays
    const [areaMasterArr, setAreaMasterArr] = React.useState<AreaMaster[]>([]);
    const [organizationMasterArr, setOrganizationMasterArr] = React.useState<OrganizationMaster[]>([]);
    const [documentsTypeMasterArr, setDocumentsTypeMasterArr] = React.useState<DocumentTypeMaster[]>([]);
    const [deliverablesMasterListArr, setDeliverablesMasterListArr] = React.useState<DeliverablesMaster[]>([]);

    // New state variables for popup and validation
    const [showPopup, setShowPopup] = React.useState(false);
    const [popupConfig, setPopupConfig] = React.useState({
        type: 'confirmation' as 'confirmation' | 'validation' | 'success' | 'error',
        title: '',
        message: ''
    });
    const [fieldErrors, setFieldErrors] = React.useState<{ [key: string]: boolean }>({});

    // Created object the handel on change in project info field
    const [projectInfo, setProjectInfo] = React.useState({
        projectName: '',
        clientName: '',
        preparedBy: '',
        startDate: undefined as Date | undefined,
        projectType: '',
        overview: ''
    });

    // Define deliverable item type
    interface DeliverableItem {
        deliverable: string;
        area: string;
        organisation: string;
        docType: string;
        seqNumber: string;
        docNumber: string;
        dueDate: Date | undefined;
        assignedTo: number;
    }

    // this is used for dynamic deleivreables table generation array
    const [deliverables, setDeliverables] = React.useState<DeliverableItem[]>([
        {
            deliverable: '',
            area: '',
            organisation: '',
            docType: '',
            seqNumber: '',
            docNumber: '',
            dueDate: undefined,
            assignedTo: 0
        }
    ]);

    // Validation function
    const validateForm = (): boolean => {
        const errors: { [key: string]: boolean } = {};

        // Validate Project Information
        if (!projectInfo.projectName.trim()) errors.projectName = true;
        if (!projectInfo.clientName.trim()) errors.clientName = true;
        if (!selectedUser) errors.preparedBy = true;
        if (!projectInfo.projectType) errors.projectType = true;
        if (!projectInfo.overview.trim()) errors.overview = true;

        // Validate Deliverables
        deliverables.forEach((deliverable, index) => {
            if (!deliverable.deliverable) errors[`deliverable_${index}`] = true;
            if (!deliverable.dueDate) errors[`dueDate_${index}`] = true;
            if (!deliverable.assignedTo) errors[`assignedTo_${index}`] = true;
        });

        setFieldErrors(errors);
        return Object.keys(errors).length === 0;
    };

    // Helper function to find user option (alternative to Array.find)
    const findUserOption = (users: UserOption[], userId: number): UserOption | null => {
        for (let i = 0; i < users.length; i++) {
            if (users[i].value === userId) {
                return users[i];
            }
        }
        return null;
    };

    // Helper function to get master data item by text value
    const getAreaByText = (areaText: string): AreaMaster | undefined => {
        return areaMasterArr.find(area => area.Area === areaText);
    };

    const getOrganizationByText = (orgText: string): OrganizationMaster | undefined => {
        return organizationMasterArr.find(org => org.Organization === orgText);
    };

    const getDocumentTypeByText = (docTypeText: string): DocumentTypeMaster | undefined => {
        return documentsTypeMasterArr.find(doc => doc.DocumentType === docTypeText);
    };

    const getDeliverableByText = (deliverableText: string): DeliverablesMaster | undefined => {
        return deliverablesMasterListArr.find(del => del.Deliverables === deliverableText);
    };

    // Sequential number and document number generation function
    const generateSequenceAndDocNumber = async (index: number, source: 'Project Type' | 'Document Type') => {
        const row = deliverables[index];
        
        if (!projectInfo.projectName || !row.area || !row.organisation || !row.docType) {
            return;
        }

        const selectedArea = getAreaByText(row.area);
        const selectedOrganization = getOrganizationByText(row.organisation);
        const selectedDocType = getDocumentTypeByText(row.docType);
        const selectedDeliverable = getDeliverableByText(row.deliverable);

        try {
            // Query to get existing items from DeliverablesDetails list
            const existingItems = await sp.web.lists.getByTitle("DeliverablesDetails")
                .items
                .select("SequenticalNumber", "ProjectCreationListID/ID", "ProjectCreationListID/ProjectName")
                .expand("ProjectCreationListID")
                .filter(`ProjectCreationListID/ProjectName eq '${projectInfo.projectName}' 
                    and Area eq '${selectedArea ? selectedArea.Area : ''}' 
                    and Organization eq '${selectedOrganization ? selectedOrganization.Organization : ''}' 
                    and DocumentType eq '${selectedDocType ? selectedDocType.DocumentType : ''}'`)
                .orderBy("ID", true)
                .top(4999)();

            // Get max sequence from SharePoint list
            const maxSeqFromList = existingItems.length > 0 
                ? Math.max(...existingItems.map(r => r.SequenticalNumber ? parseInt(r.SequenticalNumber) : 0))
                : 0;

            // Get max sequence from current deliverables array excluding current index
            const maxSeqFromArray = Math.max(...deliverables.map((r, i) => {
                if (i === index) return 0; // skip current row

                const isMatch = 
                    r.area === row.area &&
                    r.organisation === row.organisation &&
                    r.docType === row.docType;

                return isMatch && r.seqNumber ? parseInt(r.seqNumber) : 0;
            }));

            const maxSeq = Math.max(maxSeqFromList, maxSeqFromArray);

            // Increment the max sequence number
            const newSeqNumber = (maxSeq + 1).toString().padStart(4, '0');

            // Create the document number based on details
            const newDocNumber = `${projectInfo.projectName}-${
                selectedArea ? selectedArea.AreaIdentificationCode : ''
            }-${selectedArea ? selectedArea.AreaCode : ''}-${
                selectedOrganization ? selectedOrganization.OrganizationIdentificationCode : ''
            }-${selectedDocType ? selectedDocType.DocumentTypeCode : ''}-${newSeqNumber}`;

            // Update the deliverable row
            const updated = [...deliverables];
            updated[index] = {
                ...updated[index],
                seqNumber: newSeqNumber,
                docNumber: newDocNumber
            };
            setDeliverables(updated);

        } catch (error) {
            console.error("Error generating sequence and document number:", error);
        }
    };

    // Used useEffect to fetch data on mount mean on loan only
    React.useEffect(() => {
        const fetchProjectTypes = async () => {
            try {
                const items = await sp.web.lists.getByTitle("ProjectTypeMaster").items.select("Id", "ProjectType")();
                const options: IDropdownOption[] = items.map(item => ({
                    key: item.Id,
                    text: item.ProjectType,
                }));
                const placeholder: IDropdownOption = { key: '', text: 'Select Project Type' };
                setProjectTypeOptions([placeholder, ...options]);
            } catch (error) {
                console.error("Error fetching project types:", error);
            }
        };

        const fetchDeliverables = async () => {
            try {
                const items = await sp.web.lists.getByTitle("DeliverablesMaster").items.select("Id", "Deliverables")();
                const options: IDropdownOption[] = items.map(item => ({
                    key: item.Deliverables,
                    text: item.Deliverables
                }));
                const deliverablesData: DeliverablesMaster[] = items.map(item => ({
                    ID: item.Id,
                    Deliverables: item.Deliverables
                }));
                const placeholder: IDropdownOption = { key: '', text: 'Select Deliverables' };
                setDeliverablesOptions([placeholder, ...options]);
                setDeliverablesMasterListArr(deliverablesData);
            } catch (error) {
                console.error("Error fetching project types:", error);
            }
        };

        const fetchArea = async () => {
            try {
                const items = await sp.web.lists.getByTitle("AreaMaster").items.select("Id", "Area", "AreaIdentificationCode", "AreaCode")();
                const options: IDropdownOption[] = items.map(item => ({
                    key: item.Area,
                    text: item.Area
                }));
                const areaData: AreaMaster[] = items.map(item => ({
                    ID: item.Id,
                    Area: item.Area,
                    AreaIdentificationCode: item.AreaIdentificationCode,
                    AreaCode: item.AreaCode
                }));
                const placeholder: IDropdownOption = { key: '', text: 'Select Area' };
                setAreaOptions([placeholder, ...options]);
                setAreaMasterArr(areaData);
            } catch (error) {
                console.error("Error fetching project types:", error);
            }
        };

        const fetchOrganization = async () => {
            try {
                const items = await sp.web.lists.getByTitle("OrganizationMaster").items.select("Id", "Organization", "OrganizationIdentificationCode")();
                const options: IDropdownOption[] = items.map(item => ({
                    key: item.Organization,
                    text: item.Organization
                }));
                const orgData: OrganizationMaster[] = items.map(item => ({
                    ID: item.Id,
                    Organization: item.Organization,
                    OrganizationIdentificationCode: item.OrganizationIdentificationCode
                }));
                const placeholder: IDropdownOption = { key: '', text: 'Select Project Type' };
                setOrganizationOptions([placeholder, ...options]);
                setOrganizationMasterArr(orgData);
            } catch (error) {
                console.error("Error fetching Organization:", error);
            }
        };

        const fetchDocumentType = async () => {
            try {
                const items = await sp.web.lists.getByTitle("DocumentTypeMaster").items.select("Id", "DocumentType", "DocumentTypeCode")();
                const options: IDropdownOption[] = items.map(item => ({
                    key: item.DocumentType,
                    text: item.DocumentType
                }));
                const docTypeData: DocumentTypeMaster[] = items.map(item => ({
                    ID: item.Id,
                    DocumentType: item.DocumentType,
                    DocumentTypeCode: item.DocumentTypeCode
                }));
                const placeholder: IDropdownOption = { key: '', text: 'Select Document Type' };
                setDocumentTypeOptions([placeholder, ...options]);
                setDocumentsTypeMasterArr(docTypeData);
            } catch (error) {
                console.error("Error fetching DocumentType:", error);
            }
        };

        const loadUsers = async () => {
            try {
                const users = await sp.web.siteUsers();
                const options: UserOption[] = users.map((u) => ({
                    value: u.Id,
                    label: u.Title,
                    email: u.Email
                }));
                setUserOptions(options);
            } catch (err) {
                console.error("Error fetching users:", err);
            }
        };

       const setCurrentUser = async () => {
  try {
    const currentUser = await sp.web.currentUser();
    const currentUserOption: UserOption = {
      value: currentUser.Id,
      label: currentUser.Title,
      email: currentUser.Email
    };

    setSelectedUser(currentUserOption);
    handleProjectChange('preparedBy', currentUserOption.value);
  } catch (error) {
    console.error("Error fetching current user:", error);
  }
};


        loadUsers();
        fetchProjectTypes();
        fetchDeliverables();
        fetchArea();
        fetchOrganization();
        fetchDocumentType();
        setCurrentUser();
        
    }, []);

    const fetchDeliverablesByProjectType = async (projectType: string) => {
        try {
            const items = await sp.web.lists
                .getByTitle("ProjectTypeDeliverablesMaster")
                .items
                .select("Deliverables/Deliverables", "Deliverables/Id", "Area/Area", "Area/Id", "Organization/Organization", "Organization/Id", "DocumentType/DocumentType", "DocumentType/Id")
                .expand("Deliverables", "Area", "Organization", "DocumentType")
                .filter(`ProjectType/Id eq '${projectType}'`)
                .top(4999)();
            if (items.length > 0) {
                const mapped: DeliverableItem[] = await Promise.all(
                    items.map(async (item, index) => {
                        const deliverableItem: DeliverableItem = {
                            deliverable: item.Deliverables?.Deliverables || '',
                            area: item.Area?.Area || '',
                            organisation: item.Organization?.Organization || '',
                            docType: item.DocumentType?.DocumentType || '',
                            seqNumber: '',
                            docNumber: '',
                            dueDate: undefined,
                            assignedTo: 0
                        };
                        
                        // Generate sequence and doc number for each item
                        const tempDeliverables = [deliverableItem];
                        const currentIndex = 0;
                        
                        const selectedArea = getAreaByText(deliverableItem.area);
                        const selectedOrganization = getOrganizationByText(deliverableItem.organisation);
                        const selectedDocType = getDocumentTypeByText(deliverableItem.docType);

                        if (projectInfo.projectName && selectedArea && selectedOrganization && selectedDocType) {
                            try {
                                const existingItems = await sp.web.lists.getByTitle("DeliverablesDetails")
                                    .items
                                    .select("SequenticalNumber", "ProjectCreationListID/ID", "ProjectCreationListID/ProjectName")
                                    .expand("ProjectCreationListID")
                                    .filter(`ProjectCreationListID/ProjectName eq '${projectInfo.projectName}' 
                                        and Area eq '${selectedArea.Area}' 
                                        and Organization eq '${selectedOrganization.Organization}' 
                                        and DocumentType eq '${selectedDocType.DocumentType}'`)
                                    .orderBy("ID", true)
                                    .top(4999)();

                                const maxSeqFromList = existingItems.length > 0 
                                    ? Math.max(...existingItems.map(r => r.SequenticalNumber ? parseInt(r.SequenticalNumber) : 0))
                                    : 0;

                                const maxSeqFromArray = Math.max(...tempDeliverables.map((r, i) => {
                                    if (i === currentIndex) return 0;
                                    const isMatch = 
                                        r.area === deliverableItem.area &&
                                        r.organisation === deliverableItem.organisation &&
                                        r.docType === deliverableItem.docType;
                                    return isMatch && r.seqNumber ? parseInt(r.seqNumber) : 0;
                                }));

                                const maxSeq = Math.max(maxSeqFromList, maxSeqFromArray);
                                const newSeqNumber = (maxSeq + 1).toString().padStart(4, '0');
                                
                                deliverableItem.seqNumber = newSeqNumber;
                                deliverableItem.docNumber = `${projectInfo.projectName}-${
                                    selectedArea.AreaIdentificationCode
                                }-${selectedArea.AreaCode}-${
                                    selectedOrganization.OrganizationIdentificationCode
                                }-${selectedDocType.DocumentTypeCode}-${newSeqNumber}`;
                            } catch (error) {
                                console.error("Error generating sequence for pre-filled deliverables:", error);
                            }
                        }
                        
                        return deliverableItem;
                    })
                );
                setDeliverables(mapped);
            } else {
                setDeliverables([
                    {
                        deliverable: '',
                        area: '',
                        organisation: '',
                        docType: '',
                        seqNumber: '',
                        docNumber: '',
                        dueDate: undefined,
                        assignedTo: 0
                    }
                ]);
            }
        } catch (error) {
            console.error("Error fetching ProjectTypeDeliverablesMaster:", error);
            setDeliverables([
                {
                    deliverable: '',
                    area: '',
                    organisation: '',
                    docType: '',
                    seqNumber: '',
                    docNumber: '',
                    dueDate: undefined,
                    assignedTo: 0
                }
            ]);
        }
    };

    const ondocumentTypechange = async (index: number, documentType: string) => {
        try {
            const items = await sp.web.lists
                .getByTitle("DeliverablesDetails")
                .items
                .select("Deliverables", "Area", "Organization", "DocumentType")
                .filter(`DocumentType eq '${documentType}'`)
                .top(4999)();
            if (items.length > 0) {
                const mapped: DeliverableItem[] = items.map(item => ({
                    deliverable: item.Deliverables || '',
                    area: item.Area || '',
                    organisation: item.Organization || '',
                    docType: item.DocumentType || '',
                    seqNumber: '',
                    docNumber: '',
                    dueDate: undefined,
                    assignedTo: 0
                } as DeliverableItem));
                setDeliverables(mapped);
            } else {
                // Instead of replacing all deliverables, just update the current row
                const updated = [...deliverables];
                updated[index] = {
                    ...updated[index],
                    docType: documentType
                };
                setDeliverables(updated);
                
                // Generate sequence and doc number for the current row
                await generateSequenceAndDocNumber(index, 'Document Type');
            }
        } catch (error) {
            console.error("Error fetching ProjectTypeDeliverablesMaster:", error);
            // Just update the document type for the current row
            const updated = [...deliverables];
            updated[index] = {
                ...updated[index],
                docType: documentType
            };
            setDeliverables(updated);
        }
    };

    const handleCancelPopup = () => {
        setShowPopup(false);
    };

    const handleCloseValidation = () => {
        setShowPopup(false);
    };

    // Updated handleSubmit function
    const handleSubmit = () => {
        if (!validateForm()) {
            setPopupConfig({
                type: 'validation',
                title: 'Validation Error',
                message: 'Please fill all mandatory fields.'
            });
            setShowPopup(true);
            return;
        }

        setPopupConfig({
            type: 'confirmation',
            title: 'Confirm Submission',
            message: 'Are you sure you want to submit this request?'
        });
        setShowPopup(true);
    };

    // Add new function for success OK button
    const handleSuccessOk = () => {
        setShowPopup(false);
       // history.push("/Dashboard.aspx");
    };

    // Update handleConfirm to use success type
    const handleConfirm = async () => {
        setShowPopup(false);
        try {
            const projectRequestData = {
                Title: projectInfo.projectName,
                ProjectName: projectInfo.projectName,
                ProjectOverview: projectInfo.overview,
                Status: 'Pending',
                ProjectTypeId: projectInfo.projectType,
                ClientName: projectInfo.clientName,
                Date: projectInfo.startDate,
                PreparedById: projectInfo.preparedBy,
            };

            const projectRequestItem = await sp.web.lists.getByTitle("ProjectCreationList").items.add(projectRequestData);
            const projectCreationListId = projectRequestItem.data?.Id || projectRequestItem.data?.ID;

            console.log('Project request created with ID:', projectCreationListId);
            await addDeliverablesDetails(projectCreationListId);

            setPopupConfig({
                type: 'success',
                title: 'Success',
                message: 'Request submitted successfully!'
            });
            setShowPopup(true);
            window.location.reload();

        } catch (error) {
            console.error('Error submitting request:', error);

            setPopupConfig({
                type: 'error',
                title: 'Error',
                message: 'Failed to submit request. Please try again.'
            });
            setShowPopup(true);
        }
    };

    const addDeliverablesDetails = async (listId: number) => {
        try {
            for (let index = 0; index < deliverables.length; index++) {
                const row = deliverables[index];
                const assignedToId = row.assignedTo;

                // Convert sequential number to number type for SharePoint
                const sequentialNumber = row.seqNumber ? parseInt(row.seqNumber) : 0;

                const deliverableData = {
                    Title: `Deliverable ${index + 1} - ${projectInfo.projectName}`,
                    ProjectCreationListIDId: listId,
                    AssignedToId: assignedToId,
                    Status: "Pending",
                    Deliverables: row.deliverable,
                    Area: row.area,
                    Organization: row.organisation,
                    DocumentType: row.docType,
                    SequenticalNumber: Number(sequentialNumber) , // Store as number
                    DocNumber: row.docNumber, // Store as text
                    DueDate: row.dueDate,
                    RevisionNumber: "0"
                };
               const deliverableItem = await sp.web.lists.getByTitle("DeliverablesDetails").items.add(deliverableData);
                //const deliverablesId = deliverableItem.Id;
                const deliverablesId = deliverableItem.data?.Id || deliverableItem.data?.ID;
 
                await createVendorTask(deliverablesId, listId, row.docType, assignedToId);
            }
        } catch (error) {
            console.error('Error adding deliverables:', error);
            throw error;
        }
    };

    const createVendorTask = async (deliverablesId: number, listId: number, docType: string, vendorId: number) => {
        try {
            const vendorTaskData = {
                Title: `Vendor Task - ${projectInfo.projectName}`,
                ProjectCreationListIDId: listId,
                DeliverablesDetailsIdId: deliverablesId,
                AssignedToId: vendorId,
                Status: "pending",
                ApprovalCriteria: "Everyone",
                ApproverRole: "Vendor",
                SerialNumber: "0",
                Level: "Level 0",
                RequestedDate: new Date(),
                DocumentType: docType || '',
            };
            await sp.web.lists.getByTitle("ProjectApprovals").items.add(vendorTaskData);

        } catch (error) {
            console.error('Error creating vendor task:', error);
            throw error;
        }
    };

    const handleProjectChange = (field: string, value: any) => {
        setProjectInfo(prev => ({ ...prev, [field]: value }));
        console.log(projectInfo)
    };

    type DeliverableField = keyof DeliverableItem;
    const handleDeliverableChange = async (index: number, field: DeliverableField, value: string | number | Date | null | undefined) => {
        const updated = [...deliverables];
        
        if (field === 'dueDate' && value === null) {
            updated[index][field] = undefined as never;
        } else {
            updated[index][field] = value as never;
        }
        
        setDeliverables(updated);

        // Generate sequence and document number when area, organization, or document type changes
        if (field === 'area' || field === 'organisation' || field === 'docType') {
            if (field === 'area' || field === 'organisation') {
                // For area and organization changes, use 'Project Type' as source
                await generateSequenceAndDocNumber(index, 'Project Type');
            } else if (field === 'docType') {
                // For document type changes, use 'Document Type' as source
                await generateSequenceAndDocNumber(index, 'Document Type');
            }
        }
    };

    const addDeliverableRow = () => {
        setDeliverables(prev => [...prev, {
            deliverable: '',
            area: '',
            organisation: '',
            docType: '',
            seqNumber: '',
            docNumber: '',
            dueDate: undefined,
            assignedTo: 0
        }]);
    };

    const deleteDeliverableRow = (index: number) => {
        const updated = [...deliverables];
        updated.splice(index, 1);
        setDeliverables(updated);
    };

    const handleCancel = () => {
       if (props.onCancel) {
           props.onCancel();
       } else {
           // history.push("/Dashboard.aspx");
       }
    };

    return (
        <div className={styles.newRequest}>
            <h2>New Request</h2>

            {/* Project Information Section */}
            <div className={styles.section}>
                <h3 >Project Information</h3>
                {alertMessage && (
                    <div className={styles.customAlert}>
                        <span>{alertMessage}</span>
                    </div>
                )}

                <div className={styles.formRow}>
                    <TextField
                        label="Project Name"
                        placeholder="Enter Project Name"
                        required
                        value={projectInfo.projectName}
                        onChange={(_, val) => handleProjectChange('projectName', val)}
                        className={fieldErrors.projectName ? styles.textFieldError : ''}
                        
                    />
                    <TextField
                        label="Client Name"
                        placeholder="Enter Client Name"
                        required
                        value={projectInfo.clientName}
                        onChange={(_, val) => handleProjectChange('clientName', val)}
                    />
                </div>

                <div className={styles.formRow}>
                    <div className={styles.selectContainer}>
                        <label className={styles.selectLabel}>Prepared By</label>
                        <Select
                            options={userOptions}
                            value={selectedUser}
                            onChange={(option: UserOption | null) => {
                                setSelectedUser(option);
                                handleProjectChange('preparedBy', option?.value || 0); // Store UserId
                            }}
                            placeholder="Select a user..."
                            isClearable
                            required
                            className={fieldErrors.preparedBy ? styles.validationError : ''}
                            styles={{
                                control: (base: { border: any; }) => ({
                                    ...base,
                                    border: fieldErrors.preparedBy ? '1px solid #a4262c' : base.border,
                                    '&:hover': {
                                        border: fieldErrors.preparedBy ? '1px solid #a4262c' : base.border
                                    }
                                })
                            }}
                        />
                        {fieldErrors.preparedBy && (
                            <span className={styles.errorMessage}>This field is required</span>
                        )}
                    </div>
                    <DatePicker
                        label="Project Start Date"
                        placeholder="Enter Project Start Date"
                        value={projectInfo.startDate}
                        onSelectDate={(date) => handleProjectChange('startDate', date)}
                        formatDate={(date) => date ? date.toLocaleDateString('en-GB') : ''}
                    />
                </div>

                <div className={styles.formRow}>
                    <Dropdown
                        label="Project Type"
                        options={projectTypeOptions}
                        selectedKey={projectInfo.projectType}
                        onChange={async (_, option: IDropdownOption | undefined) => {
                            if (!projectInfo.projectName || projectInfo.projectName.trim() === '') {
                                setAlertMessage('⚠️ Please enter Project Name before selecting Project Type.');
                                return;
                            }
                            setAlertMessage(null);
                            const selectedType = option?.key as string;
                            handleProjectChange('projectType', selectedType)
                            await fetchDeliverablesByProjectType(selectedType);
                        }}
                        required
                    />
                    <TextField
                        label="Project Overview"
                        placeholder="Enter Project Overview"
                        required
                        multiline
                        rows={3}
                        value={projectInfo.overview}
                        onChange={(_, val) => handleProjectChange('overview', val)}
                    />
                </div>
            </div>

            {/* Deliverables Section */}
            <div className='d-flex align-items-center  justify-content-between mb-2'>
                    <h4 style={{margin:'0px'}} className='font-16 fw-bold text-dark mb-0'>Deliverables</h4>
                    <IconButton
                        iconProps={{ iconName: 'Add' }}
                        title="Add Row"
                        ariaLabel="Add Row"
                        onClick={addDeliverableRow}
                        className={styles.addButton}
                    />
                </div>
            <div className={styles.tableWrapper}>
               
                <table className={styles.table}>
                    <thead>
                        <tr>
                            <th style={{minWidth:'70px'}}>SNo</th>
                            <th>Deliverables</th>
                            <th>Area</th>
                            <th>Organisation</th>
                            <th>Document Type</th>
                            <th>Seq. No</th>
                            <th className={styles.wrapText}>Doc. No</th>
                            <th>Due Date*</th>
                            <th>Assigned To*</th>
                            <th style={{minWidth:'80px'}}>Action</th>
                        </tr>
                    </thead>
                    <tbody>
                        {deliverables.map((row, index) => (
                            <tr key={index}>
                                <td style={{minWidth:'70px'}}>{index + 1}</td>
                                <td>
                                    <Dropdown
                                        options={DeliverablesOptions}
                                        selectedKey={row.deliverable}
                                        onChange={(_: any, option: IDropdownOption | undefined) =>
                                            handleDeliverableChange(index, 'deliverable', option?.key)
                                        }
                                        placeholder="Select Deliverable"
                                    />
                                </td>
                                <td>
                                    <Dropdown
                                        options={AreaOptions}
                                        selectedKey={row.area}
                                        onChange={(_: any, option: IDropdownOption | undefined) =>
                                            handleDeliverableChange(index, 'area', option?.key)
                                        }
                                        placeholder="Select Area"
                                    />
                                </td>
                                <td>
                                    <Dropdown
                                        options={OrganizationOptions}
                                        selectedKey={row.organisation}
                                        onChange={(_: any, option: IDropdownOption | undefined) =>
                                            handleDeliverableChange(index, 'organisation', option?.key)
                                        }
                                        placeholder="Select Organisation"
                                    />
                                </td>
                                <td>
                                    <Dropdown
                                        options={DocumentTypeOptions}
                                        selectedKey={row.docType}
                                        onChange={async (_: any, option: IDropdownOption | undefined) => {
                                            const selectedDocumentType = option?.key as string;
                                            await handleDeliverableChange(index, 'docType', selectedDocumentType)
                                        }}
                                        placeholder="Select Document Type"
                                    />
                                </td>
                                <td>
                                    <TextField
                                        value={row.seqNumber}
                                        disabled
                                    />
                                </td>
                                <td>
                                    <TextField
                                        value={row.docNumber}
                                        disabled
                                    />
                                </td>
                                <td>
                                    <DatePicker
                                        value={row.dueDate}
                                        placeholder="Enter Due Date"
                                        onSelectDate={(date) => handleDeliverableChange(index, 'dueDate', date)}
                                        formatDate={(date) => date ? date.toLocaleDateString('en-GB') : ''}
                                    />
                                </td>
                                <td>
                                    <div className={styles.selectContainer}>
                                        <Select
                                            options={userOptions}
                                            value={findUserOption(userOptions, row.assignedTo)}
                                            onChange={(option: UserOption | null) =>
                                                handleDeliverableChange(index, 'assignedTo', option?.value || 0)
                                            }
                                            placeholder="Select user..."
                                            isClearable
                                        />
                                    </div>
                                </td>
                                <td style={{minWidth:'80px'}}>
                                    <IconButton
                                        iconProps={{ iconName: 'Delete' }}
                                        title="Delete Row"
                                        ariaLabel="Delete Row"
                                        onClick={() => deleteDeliverableRow(index)}
                                        className={styles.deleteButton}
                                    />
                                </td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>

            <div className={styles.buttonRow}>
                <PrimaryButton
                    text="Submit"
                    iconProps={{ iconName: 'CheckMark' }}
                    onClick={handleSubmit}
                    className={styles.submitButton}
                />
                <DefaultButton
                    text="Cancel"
                    iconProps={{ iconName: 'Cancel' }}
                    onClick={handleCancel}
                    className={styles.cancelButton}
                />
            </div>
            {/* Custom Popup */}
            <CustomPopup
                isOpen={showPopup}
                type={popupConfig.type}
                title={popupConfig.title}
                message={popupConfig.message}
                onConfirm={handleConfirm}
                onCancel={handleCancelPopup}
                onClose={handleCloseValidation}
                onSuccessOk={handleSuccessOk}
            />
        </div>
    );
};

export default NewRequest;