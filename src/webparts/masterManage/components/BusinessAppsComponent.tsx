import * as React from 'react';
// import styles from './businessappcomponent.scss';
import './businessappcomponent.css';
export const BusinessAppsComponent = ({ data }: any) => {
    React.useEffect(() => {
        // Immediately hide sidebar and navbar
        const hideElements = () => {
          document.querySelectorAll(".sidebar, .navbar , .col-lg-3").forEach((el) => {
            // (el as HTMLElement).style.display = "none";
            el.remove();
          });
        };
    
        hideElements(); // Run once when component loads
    


        // Optional: Observer to prevent SharePoint scripts from re-showing
        const observer = new MutationObserver(hideElements);
        observer.observe(document.body, { childList: true, subtree: true });
    
        return () => observer.disconnect(); // Cleanup on unmount
      }, []);
      React.useEffect(() => {
        const removeButtonsInsideIframe = () => {
          const iframe = document.getElementById("listIframe") as HTMLIFrameElement;
    
          if (!iframe || !iframe.contentDocument) return;
    
          const iframeDoc = iframe.contentDocument;
    
          // Function to remove buttons except "Add new item" and "Edit in grid view"
          const cleanUpButtons = () => {
            iframeDoc.querySelectorAll('button').forEach((button) => {
              const allowedButtons = ["new", "editInGridView"]; // Keep these buttons
              const buttonId = button.getAttribute("data-id");
    
              if (!allowedButtons.includes(buttonId)) {
                button.remove();
              }
            });
          };
    
          cleanUpButtons(); // Run once after load
    
          // MutationObserver to handle dynamic re-rendering in SharePoint
          const observer = new MutationObserver(cleanUpButtons);
          observer.observe(iframeDoc.body, { childList: true, subtree: true });
    
          // Cleanup on unmount
          return () => observer.disconnect();
        };
    
        // Run function when iframe loads
        document.getElementById("listIframe")?.addEventListener("load", removeButtonsInsideIframe);
      }, []);
    
    return <div>
        <iframe id="listIframe" src={`https://officeindia.sharepoint.com/sites/multiverseintranetportal/Lists/ARGBusinessApps?viewpath=%2Fsites%2FIntranet%2FLists%2FARGBusinessApps`}  width="100%"
      height="600px"
      style={{ border: 'none' }} />
    </div>;
};  

