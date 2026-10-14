import * as React from 'react';
import type { ICustomBannerProps } from './ICustomBannerProps';
import IBannerState from './IBannerState';


export default class CustomBanner extends React.Component<ICustomBannerProps, IBannerState> {
  constructor(props: ICustomBannerProps) {
    super(props);

    this.state = {
      userName: this.props.userDisplayName,
      greeting: "",
      isSmallScreen: window.innerWidth < 768, // initial checks
      isVerySmallScreen: window.innerWidth < 300
    };
    
    window.addEventListener('resize', this._handleResize);
  }

  public componentDidMount(): void {
    // Greeting based on local time
    const hour = new Date().getHours();
    const greeting = hour < 12 ? "Good morning" : "Good evening";
    this.setState({ greeting });
  }

  private _handleResize = (): void => {
    this.setState({
      isSmallScreen: window.innerWidth < 768,
      isVerySmallScreen: window.innerWidth < 300
    });
  }



  public render(): React.ReactElement<ICustomBannerProps> {
    return (
      <div style={{
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        fontFamily: "Segoe UI, sans-serif",
        padding: "0 10px",
        backgroundColor: this.props.backgroundColor || "#f3f2f1",
        color: this.props.fontColor || "#323130",
        border: "1px solid #ddd",
        borderRadius: "8px",
        height: "40px",
        lineHeight: "40px"
      }}>
        {/* Greeting + Display Name (Bold) */}
        <span style={{
          marginRight: "10px",
          whiteSpace: "nowrap",
          overflow: "hidden",
          textOverflow: "ellipsis"
        }}>
          {this.state.greeting}, <strong>{this.state.userName}</strong>
        </span>

        {/* Divider + Motto; adapt to screen size, hide entirely on very narrow windows */}
        {!this.state.isVerySmallScreen && (
          <>
            <span style={{
              borderLeft: "1px solid #999",
              height: "60%",
              margin: "0 10px"
            }}></span>
            <span style={{
              fontStyle: "italic",
              fontSize: this.state.isSmallScreen ? "12px" : "20px",
              textAlign: "right",
              paddingRight: "20px",
              flex: 1,
              whiteSpace: "nowrap",
              overflow: "hidden",
              textOverflow: "ellipsis"
            }}>
              “{this.props.companyMotto || "Empowering People, Driving Innovation"}”
            </span>
          </>
        )}
      </div>
    );
  }
}
