<div class="p-4 sm:p-6 bg-white rounded-xl shadow-2xl font-sans max-w-full overflow-hidden">
    <!-- Tailwind CSS CDN Load -->
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        .timeline-container {
            position: relative;
            padding: 20px 0;
            max-width: 800px;
            margin: 0 auto;
        }
        .timeline-container::after {
            content: '';
            position: absolute;
            width: 3px;
            background-color: #374151; /* Dark Gray Line */
            top: 0;
            bottom: 0;
            left: 50%;
            margin-left: -3px;
        }
        .timeline-event {
            padding: 10px 40px;
            position: relative;
            background-color: inherit;
            width: 50%;
        }
        .timeline-event::after {
            content: '';
            position: absolute;
            width: 15px;
            height: 15px;
            right: -8px;
            background-color: #10B981; /* Emerald Dot */
            border: 3px solid #1F2937; /* Dark BG Border */
            top: 15px;
            border-radius: 50%;
            z-index: 1;
        }
        .left {
            left: 0;
        }
        .right {
            left: 50%;
        }
        .left::before {
            content: " ";
            height: 0;
            position: absolute;
            top: 22px;
            width: 0;
            z-index: 1;
            right: 30px;
            border: medium solid white;
            border-width: 10px 0 10px 10px;
            border-color: transparent transparent transparent white;
        }
        .right::before {
            content: " ";
            height: 0;
            position: absolute;
            top: 22px;
            width: 0;
            z-index: 1;
            left: 30px;
            border: medium solid white;
            border-width: 10px 10px 10px 0;
            border-color: transparent white transparent transparent;
        }
        .right::after {
            left: -7px;
        }
        /* Media query for mobile view (single column) */
        @media screen and (max-width: 600px) {
            .timeline-container::after {
                left: 31px;
            }
            .timeline-event {
                width: 100%;
                padding-left: 70px;
                padding-right: 25px;
            }
            .timeline-event::before {
                left: 60px;
                border-width: 10px 10px 10px 0;
                border-color: transparent white transparent transparent;
            }
            .left::after , .right::after {
                left: 23px;
            }
            .right {
                left: 0%;
            }
        }
    </style>

    <div class="text-center mb-10">
        <h3 class="text-3xl font-extrabold text-slate-800">ESG Strategy Execution Roadmap (2022 - 2025)</h3>
        <p class="text-lg text-slate-500 mt-2">Key Milestones for Achieving Sustainable Growth & Net Zero Targets</p>
    </div>

    <!-- Timeline Content -->
    <div class="timeline-container">

        <!-- 2022: Foundation -->
        <div class="timeline-event right">
            <div class="content bg-slate-50 p-4 rounded-lg shadow-lg border-l-4 border-slate-700">
                <h4 class="text-xl font-bold text-slate-800">2022: Foundational Commitment</h4>
                <p class="text-xs text-gray-500 mb-2">Establishment & Governance</p>
                <ul class="list-disc ml-5 text-gray-700 text-sm space-y-1">
                    <li>**Governance:** Established Board-level ESG Committee and internal Steering Committee.</li>
                    <li>**Commitment:** Announced Net Zero target for operational and financed emissions.</li>
                    <li>**Measurement:** Initiated initial Scope 1 & 2 GHG emissions data collection and baseline setting.</li>
                </ul>
            </div>
        </div>

        <!-- 2023: Integration & Expansion -->
        <div class="timeline-event left">
            <div class="content bg-slate-50 p-4 rounded-lg shadow-lg border-r-4 border-slate-700">
                <h4 class="text-xl font-bold text-slate-800">2023: Risk and Strategy Integration</h4>
                <p class="text-xs text-gray-500 mb-2">Policy Development & Risk Mapping</p>
                <ul class="list-disc ml-5 text-gray-700 text-sm space-y-1">
                    <li>**Risk:** Integrated climate risk into Enterprise Risk Management (ERM) framework.</li>
                    <li>**Social:** Mandated comprehensive Ethics and Anti-Corruption training for all staff.</li>
                    <li>**Finance:** Developed Green and Sustainable Finance (GSF) Framework and sector policies.</li>
                    <li>**Board:** Achieved 100% director participation in mandatory climate-related training.</li>
                </ul>
            </div>
        </div>

        <!-- 2024: Advanced Implementation -->
        <div class="timeline-event right">
            <div class="content bg-slate-50 p-4 rounded-lg shadow-lg border-l-4 border-slate-700">
                <h4 class="text-xl font-bold text-slate-800">2024: Operational Decarbonization & Due Diligence</h4>
                <p class="text-xs text-gray-500 mb-2">Execution & New Metrics</p>
                <ul class="list-disc ml-5 text-gray-700 text-sm space-y-1">
                    <li>**Operations:** Launch of Net Zero execution plans, focusing on efficiency and renewable energy adoption.</li>
                    <li>**Supply Chain:** Implementation of Sustainable Procurement Policy and ESG supplier screening tools.</li>
                    <li>**Governance:** Linking executive compensation to key long-term ESG performance indicators.</li>
                    <li>**Data:** Completion of portfolio-wide financed emissions (Scope 3) baseline measurement.</li>
                </ul>
            </div>
        </div>

        <!-- 2025: Target Alignment & Future Outlook -->
        <div class="timeline-event left">
            <div class="content bg-slate-50 p-4 rounded-lg shadow-lg border-r-4 border-slate-700">
                <h4 class="text-xl font-bold text-slate-800">2025: Forward-Looking Compliance</h4>
                <p class="text-xs text-gray-500 mb-2">Refinement & Global Standards</p>
                <ul class="list-disc ml-5 text-gray-700 text-sm space-y-1">
                    <li>**Compliance:** Full alignment with emerging IFRS S1/S2 (ISSB) financial disclosure requirements.</li>
                    <li>**Engagement:** Formalized customer transition planning engagement strategy for high-impact clients.</li>
                    <li>**Technology:** Integration of AI Ethics and Data Governance policy into all new product development cycles.</li>
                    <li>**Review:** Conduct independent, external assurance of all major ESG disclosures and targets.</li>
                </ul>
            </div>
        </div>

    </div>

    <div class="text-center mt-8">
        <p class="text-sm text-slate-600 font-medium">➡️ Commitment extends to Net Zero Operational Emissions by 2030 and Financed Emissions by 2050.</p>
    </div>
</div>
