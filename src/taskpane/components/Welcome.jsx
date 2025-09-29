import React from "react";
import dmImage from "../../../assets/dm.png";
import { ChoiceGroup } from "@fluentui/react";

const Welcome = ({orgType, setOrgType}) => {

    const options = [
        { key: 'sandbox', text: 'Sandbox' },
        { key: 'production', text: 'Production' }
    ];

    const handleOrgTypeChange = (ev, option) => {
        if (option) {
            setOrgType(option.key);
        }
    };
    
    return (
        <div>
            <img src={dmImage} alt="DM" style={{ maxWidth: '50%', height: 'auto', margin: '0 auto', display: 'block' }} />
            <div style={{ marginTop: '20px' }}>
                <ChoiceGroup
                    selectedKey={orgType}
                    options={options}
                    onChange={handleOrgTypeChange}
                    label="Select Organization Type"
                />
            </div>
        </div>
    );
};

export default Welcome;