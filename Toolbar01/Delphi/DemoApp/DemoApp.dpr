program DemoApp;

uses
  Forms,
  Demo in 'Demo.pas' {DemoBar},
  DemoProp in 'DemoProp.pas' {PropDlg};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TDemoBar, DemoBar);
  Application.Run;
end.
