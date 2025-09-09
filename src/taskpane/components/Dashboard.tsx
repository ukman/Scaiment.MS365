import * as React from "react";

export interface DashboardProps {
    title: string;
}
  
const Dashboard: React.FC<DashboardProps> = (props: DashboardProps) => {
    const { title } = props;
    return (
        <div>
            Dashboard {title}
        </div>
      );
    
}

export default Dashboard;
