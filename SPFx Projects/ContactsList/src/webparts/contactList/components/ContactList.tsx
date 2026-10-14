import * as React from 'react';
import type { IContactListProps } from './IContactListProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IContactListState } from './IContactListState';


export default class ContactList extends React.Component<IContactListProps, IContactListState> {
  constructor(props: IContactListProps) {
    super(props);
    this.state = { items: [] };
  }
  public componentDidMount(): void {
    this._getListItems();
  }

  private _getListItems(): void {
    const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.contactListName}')/items?$select=Title,JobTitle,Company,Email,Phone,Photo`;

    this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then((data) => {
        this.setState({ items: data.value });
      });
  }


  public render(): React.ReactElement<IContactListProps> {
     return (
      <div style={{ width: "100%" }}>
      {/* Title at the top */}
      <h2 style={{
        fontFamily: "Segoe UI, sans-serif",
        fontWeight: "600",
        fontSize: "18px",
        marginBottom: "15px",
        color: "#323130"
      }}>
        Contact Directory
      </h2>

      {/* Responsive container */}
      <div style={{
        display: "flex",
        flexWrap: "wrap",
        gap: "10px"
      }}>

        {this.state.items.map(item => {
          const imageJSON = JSON.parse(item.Photo);
          console.log("Image JSON:", imageJSON);
          // const imageUrl = item.Photo ? item.Photo.Url || item.Photo.serverRelativeUrl : "";
          const imageUrl = imageJSON.serverRelativeUrl || "";

          return (
            <div key={item.Id} style={{
              display: "flex",
              alignItems: "center",
              border: "1px solid #ddd",
              borderRadius: "8px",
              padding: "10px",
              width: "300px",
              backgroundColor: "#fff",
              boxShadow: "0 2px 4px rgba(0,0,0,0.1)"
            }}>
              {/* Profile Image */}
              {imageUrl && (
                <img src={imageUrl} alt={item.Title}
                  style={{
                    width: "60px",
                    height: "60px",
                    borderRadius: "50%", // Rounded corners
                    objectFit: "cover",
                    marginRight: "15px"
                  }} />
              )}

              {/* Profile Info */}
              <div style={{ flex: 1 }}>
                <div style={{ fontWeight: "bold", fontSize: "14px" }}>{item.Title}</div>
                <div style={{ fontSize: "12px", color: "#555" }}>{item.JobTitle}</div>
                <div style={{ fontSize: "12px", color: "#555" }}>
                  Email: <a href={`mailto:${item.Email}`} style={{ color: "#0078d4", textDecoration: "none" }}>
                    {item.Email}
                  </a>
                </div>
                <div style={{ fontSize: "12px", color: "#555" }}>
                  Phone: <a href={`tel:${item.Phone}`} style={{ color: "#0078d4", textDecoration: "none" }}>
                  {item.Phone}
                </a>
              </div>
              </div>
            </div>
          );
        })}
      </div>
      </div>
    );

  }
}
