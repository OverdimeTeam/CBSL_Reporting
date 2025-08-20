"""
FastAPI application for the Automated Reporting System
"""
import logging
from typing import Dict, Any
from fastapi import FastAPI, HTTPException, Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from .routes import create_routes

logger = logging.getLogger(__name__)

def create_app(config: Dict[str, Any], scheduler) -> FastAPI:
    """Create and configure FastAPI application"""
    
    # Get API configuration
    api_config = config.get('api', {})
    
    # Create FastAPI app
    app = FastAPI(
        title=api_config.get('title', 'Automated Reporting System API'),
        version=api_config.get('version', '1.0.0'),
        description=api_config.get('description', 'API for managing automated reports'),
        docs_url="/docs",
        redoc_url="/redoc"
    )
    
    # Add CORS middleware
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],  # TODO: Configure specific origins in production
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )
    
    # Add configuration and scheduler to app state
    app.state.config = config
    app.state.scheduler = scheduler
    
    # Add exception handler
    @app.exception_handler(HTTPException)
    async def http_exception_handler(request, exc):
        logger.error(f"HTTP exception: {exc.detail}")
        return JSONResponse(
            status_code=exc.status_code,
            content={"error": exc.detail, "status_code": exc.status_code}
        )
    
    @app.exception_handler(Exception)
    async def general_exception_handler(request, exc):
        logger.error(f"Unhandled exception: {str(exc)}")
        return JSONResponse(
            status_code=500,
            content={"error": "Internal server error", "status_code": 500}
        )
    
    # Health check endpoint
    @app.get("/health")
    async def health_check():
        """Health check endpoint"""
        return {
            "status": "healthy",
            "scheduler_running": scheduler.is_running(),
            "version": api_config.get('version', '1.0.0')
        }
    
    # Root endpoint
    @app.get("/")
    async def root():
        """Root endpoint with API information"""
        return {
            "message": "Automated Reporting System API",
            "version": api_config.get('version', '1.0.0'),
            "docs": "/docs",
            "health": "/health"
        }
    
    # Startup event
    @app.on_event("startup")
    async def startup_event():
        """Application startup event"""
        logger.info("Starting Automated Reporting System API")
        
        # Start the scheduler
        try:
            scheduler.start()
            logger.info("Scheduler started successfully")
        except Exception as e:
            logger.error(f"Failed to start scheduler: {e}")
    
    # Shutdown event
    @app.on_event("shutdown")
    async def shutdown_event():
        """Application shutdown event"""
        logger.info("Shutting down Automated Reporting System API")
        
        # Stop the scheduler
        try:
            scheduler.stop()
            logger.info("Scheduler stopped successfully")
        except Exception as e:
            logger.error(f"Error stopping scheduler: {e}")
    
    # Include routes
    create_routes(app)
    
    logger.info("FastAPI application created successfully")
    return app

def get_config(request):
    """Dependency to get configuration"""
    return request.app.state.config

def get_scheduler(request):
    """Dependency to get scheduler"""
    return request.app.state.scheduler