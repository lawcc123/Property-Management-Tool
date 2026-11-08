using System.Text.Json;
using System.Windows.Controls;
namespace PropertyManager.App.PersonalizedEmail;
public partial class PersonalizedEmailView
{
    private void RestoreWorkflow(CampaignDraft saved)
    {
        // Older drafts contain the content but no navigation metadata. Infer their
        // valid completed stages, up to Review & test; never infer sending approval.
        int requestedUnlocked=Math.Clamp(saved.UnlockedStep??saved.CurrentStep??3,0,4);
        int requestedCurrent=Math.Clamp(saved.CurrentStep??requestedUnlocked,0,requestedUnlocked);
        int validUnlocked=0;string? problem=null;
        for(int step=0;step<requestedUnlocked;step++)
        {
            try{if(step<3)ValidateStage(step);validUnlocked=step+1;}
            catch(Exception ex){problem=ex.Message;break;}
        }
        int resume=problem is null?Math.Min(requestedCurrent,validUnlocked):validUnlocked;
        unlockedStep=validUnlocked;lastStep=resume;
        updatingSteps=true;
        try{Steps.SelectedIndex=resume;}
        finally{updatingSteps=false;}
        UpdateSteps();
        if(resume==2)RefreshAttachmentSizes();
        if(resume>=3)BuildPreview();
        if(resume==4)
            LiveSummary.Text="Saved draft restored. Live messages must be prepared again before sending. Choose Prepare / refresh live messages to check current sender settings and generate/validate attachments. No emails have been sent by reopening this draft.";
        string[] names=["Recipients","Message","Documents & attachments","Review & test","Send live emails"];
        StatusText.Text=problem is null
            ? "Draft restored at "+names[resume]+(resume==4?". Prepare live messages again before sending.":". Your saved content and completed steps are retained.")
            : "Draft restored. Returned to "+names[resume]+" because: "+problem;
        // Prepared campaigns, test results and send permission are deliberately not
        // restored from draft metadata. Delivery history retains actual send outcomes.
    }
}
