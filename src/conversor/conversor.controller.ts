import { Controller, Get } from '@nestjs/common';
import { ConversorService } from './conversor.service';

@Controller('conversor')
export class ConversorController {

    constructor(private service: ConversorService) {}
    @Get()
    async exportXlsx() {
        await this.service.exportXlxs();
    }
}
