"""
Report Scheduler using APScheduler
"""
import logging
from typing import Dict, Any, List
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from apscheduler.executors.pool import ThreadPoolExecutor
from .task_manager import TaskManager

logger = logging.getLogger(__name__)

class ReportScheduler:
    def __init__(self, config: Dict[str, Any]):
        self.config = config
        
        # Configure scheduler with thread pool executor
        executors = {
            'default': ThreadPoolExecutor(max_workers=3)
        }
        
        self.scheduler = BackgroundScheduler(executors=executors)
        self.task_manager = TaskManager(config)
        self._setup_jobs()
        
    def _setup_jobs(self):
        """Setup scheduled jobs based on configuration"""
        # TODO: implement job setup from config
        reports_config = self.config.get('reports', {})
        
        for report_name, report_config in reports_config.items():
            if report_config.get('enabled', False):
                schedule = report_config.get('schedule')
                if schedule:
                    self.add_scheduled_report(report_name, schedule)
                    
        logger.info(f"Setup {len(reports_config)} scheduled reports")
    
    def add_scheduled_report(self, report_name: str, cron_schedule: str):
        """Add a scheduled report job"""
        # TODO: implement scheduled job addition
        try:
            trigger = CronTrigger.from_crontab(cron_schedule)
            
            self.scheduler.add_job(
                func=self.run_report,
                trigger=trigger,
                args=[report_name],
                id=f"scheduled_{report_name}",
                name=f"Scheduled {report_name}",
                replace_existing=True,
                max_instances=1  # Prevent multiple instances of same report
            )
            
            logger.info(f"Scheduled {report_name} with cron: {cron_schedule}")
            
        except Exception as e:
            logger.error(f"Failed to schedule {report_name}: {e}")
            raise
    
    def remove_scheduled_report(self, report_name: str):
        """Remove a scheduled report job"""
        # TODO: implement job removal
        try:
            job_id = f"scheduled_{report_name}"
            self.scheduler.remove_job(job_id)
            logger.info(f"Removed scheduled report: {report_name}")
            
        except Exception as e:
            logger.error(f"Failed to remove scheduled report {report_name}: {e}")
    
    def run_report(self, report_name: str) -> str:
        """Run a report and return job ID"""
        # TODO: implement report execution
        logger.info(f"Running scheduled report: {report_name}")
        
        try:
            job_id = self.task_manager.execute_report(report_name)
            logger.info(f"Report {report_name} started with job ID: {job_id}")
            return job_id
            
        except Exception as e:
            logger.error(f"Failed to run report {report_name}: {e}")
            raise
    
    def run_report_now(self, report_name: str) -> str:
        """Run a report immediately (outside of schedule)"""
        # TODO: implement immediate report execution
        logger.info(f"Running report immediately: {report_name}")
        return self.run_report(report_name)
    
    def start(self):
        """Start the scheduler"""
        # TODO: implement scheduler start
        try:
            if not self.scheduler.running:
                self.scheduler.start()
                logger.info("Report scheduler started")
            else:
                logger.warning("Scheduler is already running")
        except Exception as e:
            logger.error(f"Failed to start scheduler: {e}")
            raise
    
    def stop(self):
        """Stop the scheduler"""
        # TODO: implement scheduler stop
        try:
            if self.scheduler.running:
                self.scheduler.shutdown(wait=True)
                logger.info("Report scheduler stopped")
            else:
                logger.warning("Scheduler is not running")
        except Exception as e:
            logger.error(f"Failed to stop scheduler: {e}")
    
    def pause_job(self, report_name: str):
        """Pause a scheduled job"""
        # TODO: implement job pausing
        try:
            job_id = f"scheduled_{report_name}"
            self.scheduler.pause_job(job_id)
            logger.info(f"Paused scheduled report: {report_name}")
        except Exception as e:
            logger.error(f"Failed to pause report {report_name}: {e}")
    
    def resume_job(self, report_name: str):
        """Resume a paused job"""
        # TODO: implement job resuming
        try:
            job_id = f"scheduled_{report_name}"
            self.scheduler.resume_job(job_id)
            logger.info(f"Resumed scheduled report: {report_name}")
        except Exception as e:
            logger.error(f"Failed to resume report {report_name}: {e}")
    
    def get_jobs(self) -> List[Dict[str, Any]]:
        """Get all scheduled jobs"""
        # TODO: implement job listing
        try:
            jobs = []
            for job in self.scheduler.get_jobs():
                jobs.append({
                    'id': job.id,
                    'name': job.name,
                    'next_run': job.next_run_time.isoformat() if job.next_run_time else None,
                    'trigger': str(job.trigger),
                    'max_instances': job.max_instances,
                    'pending': job.pending
                })
            return jobs
        except Exception as e:
            logger.error(f"Failed to get jobs: {e}")
            return []
    
    def get_job_status(self, report_name: str) -> Dict[str, Any]:
        """Get status of a specific scheduled job"""
        # TODO: implement job status retrieval
        try:
            job_id = f"scheduled_{report_name}"
            job = self.scheduler.get_job(job_id)
            
            if not job:
                return {'status': 'not_found'}
            
            return {
                'id': job.id,
                'name': job.name,
                'next_run': job.next_run_time.isoformat() if job.next_run_time else None,
                'trigger': str(job.trigger),
                'status': 'scheduled'
            }
            
        except Exception as e:
            logger.error(f"Failed to get job status for {report_name}: {e}")
            return {'status': 'error', 'error': str(e)}
    
    def is_running(self) -> bool:
        """Check if scheduler is running"""
        return self.scheduler.running
    
    def get_scheduler_info(self) -> Dict[str, Any]:
        """Get scheduler information"""
        # TODO: implement scheduler info
        return {
            'running': self.scheduler.running,
            'job_count': len(self.scheduler.get_jobs()),
            'state': self.scheduler.state,
            'timezone': str(self.scheduler.timezone) if hasattr(self.scheduler, 'timezone') else None
        }