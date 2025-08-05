
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Separator } from "@/components/ui/separator";
import { 
  Database, 
  Search, 
  FileSpreadsheet, 
  Settings, 
  Shield, 
  Users, 
  Building2,
  Trash2,
  Copy,
  Download,
  ExternalLink,
  PlayCircle,
  CheckCircle2,
  AlertTriangle,
  RefreshCw
} from "lucide-react";

const Index = () => {
  const tools = [
    {
      title: "Duplicate Detection & Merging",
      description: "Identify and merge duplicate records using built-in detection rules",
      icon: Copy,
      category: "Detection",
      difficulty: "Easy",
      steps: [
        "Configure duplicate detection rules in system settings",
        "Identify duplicates based on specific criteria",
        "Merge duplicate records to consolidate information",
        "Monitor and maintain detection rules regularly"
      ]
    },
    {
      title: "Bulk Record Deletion",
      description: "Efficiently remove obsolete records using automated deletion jobs",
      icon: Trash2,
      category: "Cleanup",
      difficulty: "Medium",
      steps: [
        "Navigate to Settings > Data Management > Bulk Record Deletion",
        "Click 'New' to create deletion job",
        "Define search criteria for target records",
        "Schedule job execution time"
      ]
    },
    {
      title: "Advanced Find & Excel Export",
      description: "Filter, analyze, and update records using Excel integration",
      icon: FileSpreadsheet,
      category: "Analysis",
      difficulty: "Easy",
      steps: [
        "Use Advanced Find to filter records",
        "Export filtered data to Excel",
        "Clean and update data in Excel",
        "Re-import cleaned data to Dynamics 365"
      ]
    },
    {
      title: "Third-Party Integration",
      description: "Enhance capabilities with specialized data cleaning tools",
      icon: Settings,
      category: "Integration",
      difficulty: "Advanced",
      steps: [
        "Evaluate third-party data cleaning solutions",
        "Ensure compatibility with your D365 version",
        "Implement automated duplicate detection",
        "Set up data enrichment services"
      ]
    },
    {
      title: "Regular Auditing",
      description: "Monitor data quality with continuous auditing practices",
      icon: Shield,
      category: "Maintenance",
      difficulty: "Medium",
      steps: [
        "Enable Dynamics 365 auditing features",
        "Review audit logs regularly",
        "Set up recurring cleanup jobs",
        "Monitor data quality metrics"
      ]
    }
  ];

  const getDifficultyColor = (difficulty: string) => {
    switch (difficulty) {
      case "Easy": return "bg-green-100 text-green-800";
      case "Medium": return "bg-yellow-100 text-yellow-800";
      case "Advanced": return "bg-red-100 text-red-800";
      default: return "bg-gray-100 text-gray-800";
    }
  };

  const getCategoryColor = (category: string) => {
    switch (category) {
      case "Detection": return "bg-blue-100 text-blue-800";
      case "Cleanup": return "bg-purple-100 text-purple-800";
      case "Analysis": return "bg-emerald-100 text-emerald-800";
      case "Integration": return "bg-orange-100 text-orange-800";
      case "Maintenance": return "bg-cyan-100 text-cyan-800";
      default: return "bg-gray-100 text-gray-800";
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 via-white to-indigo-50">
      {/* Header */}
      <div className="bg-white border-b border-gray-200 shadow-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
          <div className="flex items-center justify-between">
            <div className="flex items-center space-x-4">
              <div className="bg-blue-600 p-3 rounded-xl">
                <Database className="h-8 w-8 text-white" />
              </div>
              <div>
                <h1 className="text-3xl font-bold text-gray-900">Dynamics 365 Data Guardian</h1>
                <p className="text-gray-600 mt-1">Professional tools and methods for cleaning D365 Sales data</p>
              </div>
            </div>
            <Badge variant="secondary" className="bg-blue-100 text-blue-800 px-4 py-2">
              Enterprise Ready
            </Badge>
          </div>
        </div>
      </div>

      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {/* Overview Stats */}
        <div className="grid grid-cols-1 md:grid-cols-4 gap-6 mb-8">
          <Card className="bg-gradient-to-r from-blue-500 to-blue-600 text-white border-0">
            <CardContent className="p-6">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-blue-100 text-sm font-medium">Data Tools</p>
                  <p className="text-3xl font-bold">5</p>
                </div>
                <Database className="h-8 w-8 text-blue-200" />
              </div>
            </CardContent>
          </Card>
          
          <Card className="bg-gradient-to-r from-emerald-500 to-emerald-600 text-white border-0">
            <CardContent className="p-6">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-emerald-100 text-sm font-medium">Records</p>
                  <p className="text-3xl font-bold">Accounts & Contacts</p>
                </div>
                <Users className="h-8 w-8 text-emerald-200" />
              </div>
            </CardContent>
          </Card>

          <Card className="bg-gradient-to-r from-purple-500 to-purple-600 text-white border-0">
            <CardContent className="p-6">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-purple-100 text-sm font-medium">Integration</p>
                  <p className="text-3xl font-bold">Native & 3rd Party</p>
                </div>
                <Settings className="h-8 w-8 text-purple-200" />
              </div>
            </CardContent>
          </Card>

          <Card className="bg-gradient-to-r from-orange-500 to-orange-600 text-white border-0">
            <CardContent className="p-6">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-orange-100 text-sm font-medium">Maintenance</p>
                  <p className="text-3xl font-bold">Continuous</p>
                </div>
                <RefreshCw className="h-8 w-8 text-orange-200" />
              </div>
            </CardContent>
          </Card>
        </div>

        {/* Main Content */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          {/* Tools Grid */}
          <div className="lg:col-span-2">
            <div className="flex items-center justify-between mb-6">
              <h2 className="text-2xl font-bold text-gray-900">Data Cleaning Tools</h2>
              <div className="flex items-center space-x-2">
                <Search className="h-5 w-5 text-gray-400" />
                <span className="text-sm text-gray-500">5 tools available</span>
              </div>
            </div>

            <div className="grid gap-6">
              {tools.map((tool, index) => (
                <Card key={index} className="hover:shadow-lg transition-all duration-300 border-l-4 border-l-blue-500">
                  <CardHeader>
                    <div className="flex items-start justify-between">
                      <div className="flex items-center space-x-4">
                        <div className="bg-blue-100 p-3 rounded-lg">
                          <tool.icon className="h-6 w-6 text-blue-600" />
                        </div>
                        <div>
                          <CardTitle className="text-xl">{tool.title}</CardTitle>
                          <CardDescription className="mt-1">{tool.description}</CardDescription>
                        </div>
                      </div>
                      <div className="flex flex-col space-y-2">
                        <Badge className={getCategoryColor(tool.category)}>
                          {tool.category}
                        </Badge>
                        <Badge className={getDifficultyColor(tool.difficulty)}>
                          {tool.difficulty}
                        </Badge>
                      </div>
                    </div>
                  </CardHeader>
                  <CardContent>
                    <div className="space-y-3">
                      <h4 className="font-semibold text-gray-900 flex items-center">
                        <CheckCircle2 className="h-4 w-4 text-green-500 mr-2" />
                        Implementation Steps:
                      </h4>
                      <ul className="space-y-2">
                        {tool.steps.map((step, stepIndex) => (
                          <li key={stepIndex} className="flex items-start text-sm text-gray-600">
                            <span className="bg-blue-100 text-blue-800 rounded-full w-6 h-6 flex items-center justify-center text-xs font-medium mr-3 mt-0.5 flex-shrink-0">
                              {stepIndex + 1}
                            </span>
                            {step}
                          </li>
                        ))}
                      </ul>
                    </div>
                  </CardContent>
                </Card>
              ))}
            </div>
          </div>

          {/* Sidebar */}
          <div className="space-y-6">
            {/* Quick Actions */}
            <Card>
              <CardHeader>
                <CardTitle className="flex items-center">
                  <AlertTriangle className="h-5 w-5 text-amber-500 mr-2" />
                  Quick Actions
                </CardTitle>
              </CardHeader>
              <CardContent className="space-y-3">
                <Button className="w-full justify-start" variant="outline">
                  <Settings className="h-4 w-4 mr-2" />
                  Access D365 Settings
                </Button>
                <Button className="w-full justify-start" variant="outline">
                  <Download className="h-4 w-4 mr-2" />
                  Export Data Guide
                </Button>
                <Button className="w-full justify-start" variant="outline">
                  <Search className="h-4 w-4 mr-2" />
                  Advanced Find
                </Button>
              </CardContent>
            </Card>

            {/* Resources */}
            <Card>
              <CardHeader>
                <CardTitle>Learning Resources</CardTitle>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="bg-blue-50 p-4 rounded-lg">
                  <div className="flex items-center mb-2">
                    <PlayCircle className="h-5 w-5 text-blue-600 mr-2" />
                    <span className="font-medium text-blue-900">Video Tutorial</span>
                  </div>
                  <p className="text-sm text-blue-700 mb-3">
                    Organize and Clean your Data inside Microsoft Dynamics 365 CRM
                  </p>
                  <Button size="sm" className="bg-blue-600 hover:bg-blue-700">
                    <ExternalLink className="h-4 w-4 mr-2" />
                    Watch Now
                  </Button>
                </div>

                <Separator />

                <div className="space-y-3">
                  <h4 className="font-semibold text-gray-900">Documentation Links</h4>
                  <div className="space-y-2 text-sm">
                    <a href="#" className="flex items-center text-blue-600 hover:text-blue-700">
                      <ExternalLink className="h-3 w-3 mr-2" />
                      Microsoft Learn: Free Storage Space
                    </a>
                    <a href="#" className="flex items-center text-blue-600 hover:text-blue-700">
                      <ExternalLink className="h-3 w-3 mr-2" />
                      Stoneridge: Clean D365 CRM System
                    </a>
                    <a href="#" className="flex items-center text-blue-600 hover:text-blue-700">
                      <ExternalLink className="h-3 w-3 mr-2" />
                      AHA Apps: Audit Management
                    </a>
                  </div>
                </div>
              </CardContent>
            </Card>

            {/* Data Types */}
            <Card>
              <CardHeader>
                <CardTitle>Target Data Types</CardTitle>
              </CardHeader>
              <CardContent>
                <div className="space-y-3">
                  <div className="flex items-center justify-between p-3 bg-green-50 rounded-lg">
                    <div className="flex items-center">
                      <Building2 className="h-5 w-5 text-green-600 mr-3" />
                      <span className="font-medium text-green-900">Accounts</span>
                    </div>
                    <Badge variant="secondary" className="bg-green-100 text-green-800">
                      Primary
                    </Badge>
                  </div>
                  <div className="flex items-center justify-between p-3 bg-blue-50 rounded-lg">
                    <div className="flex items-center">
                      <Users className="h-5 w-5 text-blue-600 mr-3" />
                      <span className="font-medium text-blue-900">Contacts</span>
                    </div>
                    <Badge variant="secondary" className="bg-blue-100 text-blue-800">
                      Primary
                    </Badge>
                  </div>
                </div>
              </CardContent>
            </Card>
          </div>
        </div>
      </div>
    </div>
  );
};

export default Index;
