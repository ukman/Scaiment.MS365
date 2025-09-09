import * as React from "react";

export interface PulseProps {
    title: string;
}
  
const Pulse: React.FC<PulseProps> = (props: PulseProps) => {
    const { title } = props;
    return (
        <div>
            Pulse {title}
        </div>
      );
    
}

export default Pulse;
