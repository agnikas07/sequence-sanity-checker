import React, { useState } from "react";
import { entrypoints } from "uxp";
import { PanelController } from "./controllers/PanelController.jsx";

const app = require("premierepro");

const App = () => {
    const [status, setStatus] = useState("Ready to scan active sequence.");
    const [statusColor, setStatusColor] = useState("positive");

    const runPipelineCheck = async () => {
        try {
            const project = await app.Project.getActiveProject();
            if (!project) return triggerUpdate("Error: No active project found.", "negative");
            
            const sequence = await project.getActiveSequence();
            if (!sequence) return triggerUpdate("Error: No active sequence found.", "negative");

            const tb = await sequence.getTimebase(); 
            let fps = "Unknown";
            
            if (Number(tb) > 1000) {
                const ticks = Number(tb);
                fps = parseFloat((254016000000 / ticks).toFixed(3)).toString();
            } else {
                fps = tb.toString();
            }

            const isBroadcastSafe = (fps === "23.976" || fps === "29.97");
            const markers = sequence.markers;
            
            if (markers) {
                const newMarker = markers.createMarker(0); 
                newMarker.name = isBroadcastSafe ? "QA Passed" : "QA Failed";
                newMarker.comments = "Checked via modern UXP";
            }

            if (isBroadcastSafe) {
                triggerUpdate(`Pass: Frame rate is ${fps} fps.`, "positive");
            } else {
                triggerUpdate(`Fail: Expected 23.976 or 29.97 fps. Found: ${fps} fps`, "negative");
            }
        } catch (error) {
            triggerUpdate(`Script Error: ${error.message}`, "negative");
        }
    };

    const triggerUpdate = (message, type) => {
        setStatus(message);
        setStatusColor(type);
    };

    return (
        <div style={{ padding: "16px", display: "flex", flexDirection: "column", gap: "16px", color: "#E0E0E0", fontFamily: "sans-serif" }}>
            
            <h2 style={{ margin: 0, fontSize: "18px", fontWeight: "normal", marginBottom: "4px" }}>Pipeline QA</h2>
            
            <div style={{ 
                padding: "6px", 
                borderRadius: "6px", 
                backgroundColor: "#2D2D2D",
                borderLeft: statusColor === "positive" ? "6px solid #23A330" : "6px solid #E81A25",
                display: "flex",
                flexDirection: "column",
                gap: "8px"
            }}>
                <span style={{ fontSize: "14px", color: "#999", letterSpacing: "1px", textTransform: "uppercase", marginBottom: "6px" }}>
                    <strong>Status</strong>
                </span>
                <span style={{ fontSize: "18px", lineHeight: "1.4", color: "#FFF" }}>
                    {status}
                </span>
            </div>

            <button uxp-variant="cta" onClick={runPipelineCheck}>
                Run Sequence Check
            </button>
            
        </div>
    );
};

const myController = new PanelController(() => <App />, { 
    id: "demos" 
});
entrypoints.setup({ 
    panels: { 
        demos: myController 
    } 
});